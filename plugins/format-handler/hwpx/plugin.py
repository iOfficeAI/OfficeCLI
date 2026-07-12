#!/usr/bin/env python3
"""OfficeCLI HWPX Format Handler Plugin (Python, protocol v1)."""

import sys
import json
import io
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph'
HS = 'http://www.hancom.co.kr/hwpml/2011/section'
OPF = 'http://www.idpf.org/2007/opf'


def _emit(obj: Dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + '\n')
    sys.stdout.flush()


def _result(value: Any = None) -> Dict[str, Any]:
    return {'protocol': 1, 'msg_type': 'ok', 'result': value}


def _error(code: str, message: str, detail: str = '') -> Dict[str, Any]:
    return {
        'protocol': 1,
        'msg_type': 'error',
        'error': {'code': code, 'message': message, 'detail': detail},
    }


class HwpxDocument:
    def __init__(self, path: str):
        self.path = Path(path)
        self._raw_bytes = self.path.read_bytes()
        self.zip = zipfile.ZipFile(io.BytesIO(self._raw_bytes), 'r')
        self._sections: List[ET.Element] = []
        self._load()

    def _load(self) -> None:
        content_data = self._read_content_hpf()
        content_root = ET.fromstring(content_data)
        section_files = self._collect_section_files(content_root)
        if not section_files:
            section_files = ['Contents/section0.xml']
        for sec_file in section_files:
            try:
                sec_data = self.zip.read(sec_file)
                self._sections.append(ET.fromstring(sec_data))
            except KeyError:
                continue
        if not self._sections:
            self._sections = [
                ET.fromstring(f'<hs:sec xmlns:hp="{HP}" xmlns:hs="{HS}"></hs:sec>')
            ]

    def _read_content_hpf(self) -> bytes:
        for name in ['Contents/content.hpf', 'content.hpf', 'META-INF/manifest.xml']:
            try:
                return self.zip.read(name)
            except KeyError:
                continue
        raise KeyError('HWPX content manifest not found')

    def _collect_section_files(self, content_root: ET.Element) -> List[str]:
        section_files: List[str] = []
        for item in content_root.iter(f'{{{OPF}}}item'):
            href = item.get('href', '')
            if 'section' in href.lower() and href.endswith('.xml'):
                section_files.append(href)
        return section_files

    @property
    def section_count(self) -> int:
        return len(self._sections)

    def _collect_text(self, element: ET.Element) -> str:
        return ''.join((t.text or '') for t in element.iter(f'{{{HP}}}t'))

    def get_paragraphs(
        self, section_idx: int = 0
    ) -> List[Dict[str, Any]]:
        if section_idx >= len(self._sections):
            return []
        section = self._sections[section_idx]
        result: List[Dict[str, Any]] = []
        for i, p in enumerate(section.iter(f'{{{HP}}}p')):
            result.append(
                {
                    'tag': 'paragraph',
                    'id': p.get('id', ''),
                    'text': self._collect_text(p),
                    'style': p.get('styleIDRef', ''),
                    'path': f'/section[{section_idx}]/paragraph[{i}]',
                }
            )
        return result

    def get_node(self, path: str) -> Dict[str, Any]:
        parts = path.strip('/').split('/')
        section_idx = 0
        target_idx = None
        target_tag = None
        for part in parts:
            if part.startswith('section['):
                section_idx = int(part.split('[')[1].split(']')[0])
            elif part.startswith('paragraph['):
                target_idx = int(part.split('[')[1].split(']')[0])
                target_tag = 'paragraph'
        if target_tag == 'paragraph':
            paras = self.get_paragraphs(section_idx)
            if target_idx < len(paras):
                return paras[target_idx]
            raise KeyError('paragraph not found')
        return {
            'tag': 'section',
            'id': str(section_idx),
            'path': f'/section[{section_idx}]',
            'section_index': section_idx,
        }

    def set_paragraph_text(
        self, section_idx: int, para_idx: int, text: str
    ) -> bool:
        if section_idx >= len(self._sections):
            return False
        paras = list(self._sections[section_idx].iter(f'{{{HP}}}p'))
        if para_idx >= len(paras):
            return False
        para = paras[para_idx]
        for run in list(para.iter(f'{{{HP}}}run')):
            para.remove(run)
        run = ET.SubElement(para, f'{{{HP}}}run')
        t = ET.SubElement(run, f'{{{HP}}}t')
        t.text = text
        return True

    def add_paragraph(
        self,
        section_idx: int,
        text: str,
        style: str = '',
        position: Optional[Dict[str, Any]] = None,
    ) -> Optional[str]:
        if section_idx >= len(self._sections):
            return None
        section = self._sections[section_idx]
        para = ET.Element(f'{{{HP}}}p')
        if style:
            para.set('styleIDRef', style)
        try:
            next_id = max(
                int(p.get('id', '0')) for p in section.iter(f'{{{HP}}}p')
            )
        except ValueError:
            next_id = -1
        para.set('id', str(next_id + 1))
        run = ET.SubElement(para, f'{{{HP}}}run')
        t = ET.SubElement(run, f'{{{HP}}}t')
        t.text = text

        if position and 'index' in position:
            children = list(section)
            idx = int(position['index'])
            if idx < len(children):
                section.insert(idx, para)
            else:
                section.append(para)
        else:
            section.append(para)
        return para.get('id')

    def remove_element(
        self, section_idx: int, para_idx: int
    ) -> bool:
        if section_idx >= len(self._sections):
            return False
        paras = list(self._sections[section_idx].iter(f'{{{HP}}}p'))
        if para_idx >= len(paras):
            return False
        self._sections[section_idx].remove(paras[para_idx])
        return True

    def save(self, output_path: Optional[str] = None) -> None:
        if output_path is None:
            output_path = str(self.path)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for item in self.zip.namelist():
                data = self.zip.read(item)
                if item.lower().endswith('.xml') and 'section' in item.lower():
                    data = self._encode_section(item)
                zf.writestr(item, data)

    def _encode_section(self, href: str) -> bytes:
        try:
            if 'section' in href.lower() and href.lower().endswith('.xml'):
                idx = int(''.join(ch for ch in href if ch.isdigit()))
                if 0 <= idx < len(self._sections):
                    return ET.tostring(self._sections[idx], encoding='unicode').encode(
                        'utf-8'
                    )
        except Exception:
            pass
        return self.zip.read(href)

    def close(self) -> None:
        self.zip.close()


def _doc() -> HwpxDocument:
    doc = globals().get('_doc_instance')
    if doc is None:
        raise RuntimeError('No document open')
    return doc


def handle_get(doc: HwpxDocument, args: Dict[str, Any]) -> Dict[str, Any]:
    path = args.get('path', '/')
    try:
        return _result(doc.get_node(path))
    except Exception as exc:
        return _error('not_found', str(exc))


def handle_set(
    doc: HwpxDocument, args: Dict[str, Any], props: Dict[str, Any]
) -> Dict[str, Any]:
    path = args.get('path', '')
    parts = path.strip('/').split('/')
    section_idx = 0
    para_idx = None
    for part in parts:
        if part.startswith('section['):
            section_idx = int(part.split('[')[1].split(']')[0])
        elif part.startswith('paragraph['):
            para_idx = int(part.split('[')[1].split(']')[0])
    unsupported = [k for k in props.keys() if k != 'text']
    if para_idx is not None and 'text' in props:
        if not doc.set_paragraph_text(section_idx, para_idx, props['text']):
            return _error('invalid_argument', 'failed to set text')
    return _result({'unsupported_properties': unsupported})


def handle_add(
    doc: HwpxDocument, args: Dict[str, Any], props: Dict[str, Any]
) -> Dict[str, Any]:
    parent = args.get('parent_path', args.get('parent', '/'))
    add_type = args.get('type', 'paragraph')
    position = args.get('position')
    parts = parent.strip('/').split('/')
    section_idx = 0
    for part in parts:
        if part.startswith('section['):
            section_idx = int(part.split('[')[1].split(']')[0])

    if add_type == 'paragraph':
        new_id = doc.add_paragraph(
            section_idx,
            props.get('text', ''),
            props.get('style', ''),
            position,
        )
        if new_id is not None:
            return _result(
                {
                    'path': f'/section[{section_idx}]/paragraph[{new_id}]',
                    'unsupported_properties': [],
                }
            )
        return _error('invalid_argument', 'failed to add paragraph')

    return _error('unsupported_feature', f'Unsupported type: {add_type}')


def handle_query(
    doc: HwpxDocument, args: Dict[str, Any]
) -> Dict[str, Any]:
    selector = (args.get('selector') or '').lower()
    results: List[Dict[str, Any]] = []
    for sec_idx in range(doc.section_count):
        for para in doc.get_paragraphs(sec_idx):
            text = para.get('text', '')
            if not selector or selector in text.lower():
                results.append(para)
    return _result(results)


def handle_remove(
    doc: HwpxDocument, args: Dict[str, Any]
) -> Dict[str, Any]:
    path = args.get('path', '')
    parts = path.strip('/').split('/')
    section_idx = 0
    para_idx = None
    for part in parts:
        if part.startswith('section['):
            section_idx = int(part.split('[')[1].split(']')[0])
        elif part.startswith('paragraph['):
            para_idx = int(part.split('[')[1].split(']')[0])
    if para_idx is not None:
        return _result(
            'removed' if doc.remove_element(section_idx, para_idx) else None
        )
    return _result(None)


def handle_info() -> Dict[str, Any]:
    return {
        'name': 'officecli-hwpx-py',
        'version': '0.1.0',
        'protocol': 1,
        'kinds': ['format-handler'],
        'extensions': ['.hwpx'],
        'target': 'docx',
        'runtime': 'python',
        'idle_timeout_seconds': {
            'default': 30,
            'verbs': {'save': 60, 'dump': 30},
        },
        'vocabulary': {
            'addable_types': ['paragraph', 'section'],
            'settable_props': {'paragraph': ['text', 'style']},
            'path_segments': ['/section[N]', '/section[N]/paragraph[M]'],
        },
        'description': 'Minimal HWPX format handler',
        'supports': ['paragraphs'],
    }


def main() -> None:
    doc: Optional[HwpxDocument] = None
    for raw in sys.stdin:
        line = raw.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            _emit(_error('invalid_request', 'Invalid JSON on stdin'))
            continue

        protocol = msg.get('protocol')
        msg_type = msg.get('msg_type', '')
        if protocol != 1:
            _emit(
                _error(
                    'protocol_mismatch',
                    f'Expected protocol 1, got {protocol}',
                )
            )
            continue

        if msg_type == 'open':
            doc = HwpxDocument(msg.get('path', ''))
            snapshot = {
                'capabilities': {
                    'commands': [
                        'add',
                        'set',
                        'get',
                        'query',
                        'remove',
                        'save',
                        'view',
                        'raw',
                    ],
                    'features': ['save'],
                },
                'vocabulary': {
                    'addable_types': ['paragraph', 'section'],
                    'settable_props': {'paragraph': ['text', 'style']},
                    'path_segments': [
                        '/section[N]',
                        '/section[N]/paragraph[M]',
                    ],
                },
            }
            _emit(_result(snapshot))
            continue

        if msg_type == 'command':
            if doc is None:
                _emit(
                    _error(
                        'plugin_stream_closed', 'No document open'
                    )
                )
                continue
            command = msg.get('command', '')
            args = msg.get('args', {})
            props = msg.get('props', {})
            if command == 'get':
                _emit(handle_get(doc, args))
            elif command == 'set':
                _emit(handle_set(doc, args, props))
            elif command == 'add':
                _emit(handle_add(doc, args, props))
            elif command == 'query':
                _emit(handle_query(doc, args))
            elif command == 'remove':
                _emit(handle_remove(doc, args))
            else:
                _emit(
                    _error(
                        'unsupported_command',
                        f'Unsupported command: {command}',
                    )
                )
            continue

        if msg_type == 'ping':
            _emit(_result(None))
            continue

        if msg_type == 'save':
            if doc:
                doc.save()
            _emit(_result(None))
            continue

        if msg_type == 'close':
            if doc:
                doc.close()
                doc = None
            _emit(_result(None))
            return

        _emit(_error('invalid_request', f'Unknown msg_type: {msg_type}'))


if __name__ == '__main__':
    main()
