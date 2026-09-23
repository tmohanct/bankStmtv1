"""Classify explicit AI metadata separately from proven statement edits."""
from __future__ import annotations

from dataclasses import dataclass
import re
from xml.etree import ElementTree as ET

from src.utils.pdf_provenance_reader import read_c2pa

AI_TOOL = re.compile(
    r'\b(?:ChatGPT|OpenAI|GPT-\d[\w.-]*|DALL[ -]?E(?:\s*\d)?|Midjourney|'
    r'Stable Diffusion|Adobe Firefly|Google Gemini|Anthropic Claude)\b', re.I)
AI_SOURCE_TYPES = {'trainedAlgorithmicMedia', 'compositeWithTrainedAlgorithmicMedia'}
LIMITATION = ('Metadata records AI processing, not which statement values changed. '
              'C2PA signatures, assertion bindings and issuer trust are not validated by this check. '
              'Metadata can be removed or forged; no evidence does not mean an untouched bank original.')


@dataclass(frozen=True)
class AIProvenanceResult:
    code: str
    result: str
    details: str
    detected: bool = False
    incomplete: bool = False

    @property
    def status(self):
        return 'WARNING' if self.detected or self.incomplete else 'PASS'


def _text(value):
    return value[:500] if isinstance(value, str) else ''


def _agent(value):
    return _text(value.get('name')) if isinstance(value, dict) else _text(value)


def _ai_source(value):
    return _text(value).rstrip('/').rsplit('/', 1)[-1] in AI_SOURCE_TYPES


def inspect_ai_provenance(document) -> AIProvenanceResult:
    """Return stable machine-readable status and workbook-ready evidence.

    Only dedicated software/source/action metadata is classified. Neither bank
    statement text nor arbitrary attachment text is searched for AI keywords.
    """
    findings, metadata_findings, errors = [], [], []
    present = False
    try:
        records, errors, present = read_c2pa(document)
        for path, record in records:
            if path.endswith(('c2pa.claim', 'c2pa.claim.v2')):
                agents = record.get('claim_generator_info', [])
                if isinstance(agents, dict):
                    agents = [agents]
                if not isinstance(agents, list):
                    raise ValueError('Invalid C2PA claim generator list')
                names = [_agent(item) for item in agents]
                names.append(_text(record.get('claim_generator')))
                for name in filter(None, names):
                    if AI_TOOL.search(name):
                        findings.append(f'{path}: claim generator={name}')
            else:
                actions = record.get('actions', [])
                if not isinstance(actions, list):
                    raise ValueError('Invalid C2PA actions list')
                for action in actions:
                    if not isinstance(action, dict):
                        raise ValueError('Invalid C2PA action')
                    agent = _agent(action.get('softwareAgent'))
                    source = _text(action.get('digitalSourceType'))
                    # A softwareAgentIndex refers to the assertion's agent list.
                    agent_index = action.get('softwareAgentIndex')
                    agents = record.get('softwareAgents', [])
                    if type(agent_index) is int and isinstance(agents, list) and 0 <= agent_index < len(agents):
                        agent = _agent(agents[agent_index])
                    if AI_TOOL.search(agent) or _ai_source(source):
                        findings.append(f'{path}: action={_text(action.get("action")) or "unspecified"}; '
                                        f'software={agent or "unspecified"}; source={source or "unspecified"}; '
                                        f'when={_text(action.get("when")) or "not recorded"}')
                if record.get('allActionsIncluded') is False:
                    findings_context = 'C2PA declares that its action history is incomplete.'
                    # Context alone is never evidence of AI use.
                    if findings:
                        findings.append(findings_context)
    except Exception as exc:
        errors.append(f'C2PA inspection incomplete: {type(exc).__name__}: {exc}')
    try:
        for key in ('creator', 'producer'):
            value = _text((document.metadata or {}).get(key))
            if AI_TOOL.search(value):
                metadata_findings.append(f'PDF {key}: {value}')
        xmp = document.get_xml_metadata() or ''
        if xmp:
            for element in ET.fromstring(xmp).iter():
                for key, value in [(element.tag, element.text), *element.attrib.items()]:
                    # Only standard XMP/IPTC fields; document titles aren't signals.
                    if key in ('{http://ns.adobe.com/xap/1.0/}CreatorTool',
                               '{http://ns.adobe.com/xap/1.0/sType/ResourceEvent#}softwareAgent'):
                        if AI_TOOL.search(_text(value)):
                            metadata_findings.append(f'XMP software: {_text(value)}')
                    elif key == '{http://iptc.org/std/Iptc4xmpExt/2008-02-29/}DigitalSourceType':
                        source = _text(value) or _text(element.attrib.get('{http://www.w3.org/1999/02/22-rdf-syntax-ns#}resource'))
                        if _ai_source(source):
                            metadata_findings.append(f'XMP DigitalSourceType: {source}')
    except Exception as exc:
        errors.append(f'AI metadata inspection incomplete: {type(exc).__name__}: {exc}')
    if findings:
        code, result = 'AI_PROVENANCE_RECORDED', 'AI processing recorded in Content Credentials'
    elif metadata_findings:
        code, result = 'AI_METADATA_RECORDED', 'AI software/source recorded in unsigned metadata'
    elif errors:
        code, result = 'INCONCLUSIVE', 'AI provenance inspection incomplete'
    else:
        code, result = 'NO_AI_EVIDENCE', 'No AI evidence found'
    details = [f'AI status: {code}', *list(dict.fromkeys(findings + metadata_findings))[:30]]
    if present and not findings:
        details.append('Content Credentials present; no supported explicit AI declaration was found.')
    details.extend(errors[:10])
    details.append(LIMITATION)
    return AIProvenanceResult(code, result, '\n'.join(details)[:12000],
                              bool(findings or metadata_findings), bool(errors))
