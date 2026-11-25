"""Load private conversion rules from the uploaded client workbook.

This module contains no inventory data. The optional deployment setting
SMILING_ROCKS_RULES_FILE can point to a private zlib JSON file. Normal operation
reads the hidden _VDB_Rules sheet in the supplied client workbook/template.
"""
import base64
import io
import json
import os
import shutil
import tempfile
from pathlib import PurePosixPath
import zipfile
import zlib
import xml.etree.ElementTree as E

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def _read_workbook(source):
    with zipfile.ZipFile(source) as archive:
        if 'xl/workbook.xml' not in archive.namelist():
            for name in archive.namelist():
                if name.lower().endswith('.xlsx') and not name.startswith('__MACOSX/'):
                    with archive.open(name) as original,tempfile.TemporaryFile() as nested:
                        shutil.copyfileobj(original,nested);nested.seek(0)
                        rules=_read_workbook(nested)
                    if rules:return rules
            return None
        book=E.fromstring(archive.read('xl/workbook.xml'))
        sheet=next((s for s in book.findall('s:sheets/s:sheet',NS) if s.get('name')=='_VDB_Rules'),None)
        if sheet is None:return None
        rid=sheet.get('{'+NS['r']+'}id')
        relationships=E.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        target=next(r.get('Target') for r in relationships if r.get('Id')==rid)
        if target.startswith('/'):path=target.lstrip('/')
        else:path=str(PurePosixPath('xl')/target)
        if '..' in PurePosixPath(path).parts:raise ValueError('Invalid saved-rule worksheet path.')
        root=E.fromstring(archive.read(path));shared=[]
        if 'xl/sharedStrings.xml' in archive.namelist():
            shared=[''.join(si.itertext()) for si in E.fromstring(archive.read('xl/sharedStrings.xml'))]
        values={}
        for row in root.findall('s:sheetData/s:row',NS):
            for c in row:
                if c.get('t')=='inlineStr':value=''.join(c.find('s:is',NS).itertext())
                elif c.get('t')=='s':value=shared[int(c.findtext('s:v','0',NS))]
                else:value=c.findtext('s:v','',NS)
                values[c.get('r')]=value
        if values.get('A1')!='VDB_RULES_V1':raise ValueError('Unrecognized saved conversion rules version.')
        chunks=[(int(key[1:]),value) for key,value in values.items() if key.startswith('A') and key[1:].isdigit() and int(key[1:])>=4]
        encoded=''.join(value for _,value in sorted(chunks))
        return json.loads(zlib.decompress(base64.b64decode(encoded)))


def load_rules(sources=None):
    found=None
    for source in sources or []:
        rules=_read_workbook(source)
        if rules is not None:
            if found is not None and found!=rules:raise ValueError('Uploaded workbooks contain different saved rule versions.')
            found=rules
    if found is not None:return found
    private_path=os.environ.get('SMILING_ROCKS_RULES_FILE')
    if private_path:
        with open(private_path,'rb') as f:return json.loads(zlib.decompress(f.read()))
    return {'profiles':{},'source_labels':{}}
