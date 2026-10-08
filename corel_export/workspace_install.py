"""Merge our toolbar into a closed Corel workspace without resetting it."""
from pathlib import Path
import io
import re
import zipfile
import xml.etree.ElementTree as ET

BAR = 'a6c07fb3-ef6e-44af-89b8-efab393015ce'
EXPORT = '87260fe1-e755-4c95-9f68-2aa2f3dc5001'
HELP = '87260fe1-e755-4c95-9f68-2aa2f3dc5002'
CHECK = '87260fe1-e755-4c95-9f68-2aa2f3dc5003'
JOIN = '87260fe1-e755-4c95-9f68-2aa2f3dc5004'
BUTTONS = ((EXPORT, 'DXF', 'ShowExporter', 'CorelDXF.ico'),
           (CHECK, 'Проверить', 'ShowValidator', 'VectorCheck.ico'))
STANDARD = 'c2b44f69-6dec-444e-a37e-5dbf7ff43dae'
CATEGORY = '2cc24a3e-fe24-4708-9a74-9c75406eebcd'


def merge_workspace(data, icons):
    source = zipfile.ZipFile(io.BytesIO(data))
    parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True))
    root = ET.fromstring(source.read('content/workspace.xml'), parser=parser)
    items = root.find('items')
    bars = root.find('commandBars')
    if items is None or bars is None:
        raise ValueError('Unsupported workspace structure')
    old_items = {x.get('guid') for x in items
                 if 'SkladCorelDXF' in x.get('dynamicCommand', '')}
    standard = bars.find(f'commandBarData[@guid="{STANDARD}"]/toolbar')
    if standard is None:
        raise ValueError('Workspace has no standard toolbar anchor')
    # Keep the original XML formatting and unrelated nodes byte-for-byte.
    xml = source.read('content/workspace.xml').decode('utf-8')
    # Corel may regenerate command GUIDs after a macro reload. Keep user
    # shortcuts connected when normalizing our two installed command IDs.
    for item in items:
        command = item.get('dynamicCommand', '')
        mapped = {'SkladCorelDXF.ExportDXF.'+macro: guid for guid,caption,macro,icon in BUTTONS}.get(command)
        old = item.get('guid')
        if mapped and old and old != mapped:
            xml = re.sub(r'(<keySequence\b[^>]*\bitemRef=")' + re.escape(old) + r'(")',
                         lambda match: match[1] + mapped + match[2], xml)
    for guid in (old_items | {JOIN, HELP} | {x[0] for x in BUTTONS}) - {None}:
        escaped = re.escape(guid)
        xml = re.sub(r'<itemData\b(?=[^>]*\bguid="' + escaped + r'")[^>]*?(?:/>|>.*?</itemData>)', '', xml, flags=re.S)
        xml = re.sub(r'<item\b(?=[^>]*\bguidRef="' + escaped + r'")[^>]*?(?:/>|>.*?</item>)', '', xml, flags=re.S)
    definitions = ''.join(
        '<itemData guid="' + guid + '" dynamicCategory="' + CATEGORY
        + '" dynamicCommand="SkladCorelDXF.ExportDXF.' + macro
        + '" userCaption="' + caption + '" icon="guid://' + guid + '" enable="true"></itemData>'
        for guid, caption, macro, icon in BUTTONS)
    xml = xml.replace('</items>', definitions + '</items>', 1)
    refs = ''.join('<item guidRef="' + guid + '" itemFace="textRightOfImage"></item>' for guid,caption,macro,icon in BUTTONS)
    # A real, independently movable toolbar, not buttons on Standard.
    xml = re.sub(r'<commandBarData\b(?=[^>]*\bguid="' + BAR
                 + r'")[^>]*?(?:/>|>.*?</commandBarData>)', '', xml, flags=re.S)
    definition = ('<commandBarData guid="' + BAR + '" type="toolbar" '
                  'nonLocalizableName="Sklad DXF" userCaption="Sklad DXF" userCreated="true">'
                  '<toolbar type="toolbar" dock="fill" itemFace="textRightOfImage">' + refs + '</toolbar></commandBarData>')
    xml = xml.replace('</commandBars>', definition + '</commandBars>', 1)
    if not root.findall(f'.//toolbar[@guidRef="{BAR}"]'):
        pattern = r'(<toolbar\b(?=[^>]*\bguidRef="' + STANDARD + r'")[^>]*?(?:/>|>\s*</toolbar>))'
        xml, count = re.subn(pattern, lambda match: match.group(1)
                            + '<toolbar guidRef="' + BAR + '" dock="top" userVisibility="true"></toolbar>', xml)
        if not count:
            raise ValueError('Cannot locate standard toolbar docking anchor')
    ET.fromstring(xml)
    replacements = {'content/workspace.xml': xml.encode('utf-8')}
    for guid,caption,macro,filename in BUTTONS:
        replacements['content/icons/' + guid + '.ico'] = (Path(icons) / filename).read_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as target:
        for entry in source.infolist():
            target.writestr(entry, replacements.pop(entry.filename, source.read(entry.filename)))
        for name, value in replacements.items():
            target.writestr(name, value)
    return out.getvalue()
