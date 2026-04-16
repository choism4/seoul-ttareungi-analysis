#!/usr/bin/env python3
"""OOXML 피벗테이블 + 스파크라인 주입"""
import json, os, shutil, zipfile, re
from lxml import etree

os.chdir(os.path.dirname(os.path.abspath(__file__)))

XLSX = '따릉이_이용패턴_분석.xlsx'
with open('excel_meta.json') as f:
    meta = json.load(f)
with open('age_yearly.json') as f:
    age_yearly = json.load(f)

years = meta['years']
age_groups = meta['age_groups']
SPARK_START = meta['SPARK_START']
SPARK_END = meta['SPARK_END']

# ── Namespace map ──
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
    'xm': 'http://schemas.microsoft.com/office/excel/2006/main',
}

# Copy xlsx
TMP = XLSX + '.tmp'
shutil.copy2(XLSX, TMP)

# ══ PIVOT TABLE ══
# Create: pivotCacheDefinition, pivotCacheRecords, pivotTable
# Source: Pivot_Source sheet, A3:I10 (age groups × years)

# Prepare pivot data
# Rows = age groups, Cols = years, Values = usage counts
pivot_records = []
for ay in age_yearly:
    for ag in age_groups:
        pivot_records.append({
            '연도': ay['연도'],
            '연령대': ag,
            '이용건수': ay.get(ag, 0)
        })

unique_years = sorted(set(r['연도'] for r in pivot_records))
unique_ages = age_groups[:]

# pivotCacheDefinition1.xml
cache_def = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  r:id="rId1" refreshOnLoad="0" createdVersion="8" refreshedVersion="8" recordCount="{len(pivot_records)}">
  <cacheSource type="worksheet">
    <worksheetSource ref="A3:C{3+len(pivot_records)}" sheet="Pivot_Source"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="연도" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="{min(unique_years)}" maxValue="{max(unique_years)}" count="{len(unique_years)}">
        {chr(10).join(f'        <n v="{y}"/>' for y in unique_years)}
      </sharedItems>
    </cacheField>
    <cacheField name="연령대" numFmtId="0">
      <sharedItems count="{len(unique_ages)}">
        {chr(10).join(f'        <s v="{a}"/>' for a in unique_ages)}
      </sharedItems>
    </cacheField>
    <cacheField name="이용건수" numFmtId="3">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="0" maxValue="99999999"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>'''

# pivotCacheRecords1.xml
records_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  count="{len(pivot_records)}">
'''
for rec in pivot_records:
    yr_idx = unique_years.index(rec['연도'])
    age_idx = unique_ages.index(rec['연령대'])
    records_xml += f'  <r><x v="{yr_idx}"/><x v="{age_idx}"/><n v="{rec["이용건수"]}"/></r>\n'
records_xml += '</pivotCacheRecords>'

# pivotTable1.xml
# Rows: 연령대 (field 1), Cols: 연도 (field 0), Data: 이용건수 sum (field 2)
row_items = ''
for i in range(len(unique_ages)):
    row_items += f'    <i><x v="{i}"/></i>\n'
row_items += '    <i t="grand"><x/></i>\n'

col_items = ''
for i in range(len(unique_years)):
    col_items += f'    <i><x v="{i}"/></i>\n'
col_items += '    <i t="grand"><x/></i>\n'

pivot_table = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotTable1" cacheId="0" applyNumberFormats="0" applyBorderFormats="0"
  applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0"
  applyWidthHeightFormats="1" dataCaption="합계" updatedVersion="8" minRefreshableVersion="3"
  useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1"
  outlineData="1" multipleFieldFilters="0">
  <location ref="A3:{chr(65+len(unique_years)+1)}{4+len(unique_ages)}"
    firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisCol" showAll="0">
      <items count="{len(unique_years)+1}">
        {chr(10).join(f'        <item x="{i}"/>' for i in range(len(unique_years)))}
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisRow" showAll="0">
      <items count="{len(unique_ages)+1}">
        {chr(10).join(f'        <item x="{i}"/>' for i in range(len(unique_ages)))}
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField dataField="1" numFmtId="3" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="1"/></rowFields>
  <rowItems count="{len(unique_ages)+1}">
{row_items}  </rowItems>
  <colFields count="1"><field x="0"/></colFields>
  <colItems count="{len(unique_years)+1}">
{col_items}  </colItems>
  <dataFields count="1">
    <dataField name="합계 : 이용건수" fld="2" baseField="1" baseItem="0" numFmtId="3"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1"
    showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>'''

# ── Write into xlsx zip ──
# We need to:
# 1. Add pivot cache + table XML files
# 2. Update [Content_Types].xml
# 3. Add relationships
# 4. Create a new "Pivot_Table" sheet that references the pivot

# First, let's modify the xlsx
import tempfile

with zipfile.ZipFile(TMP, 'r') as zin:
    existing_files = zin.namelist()

    tmpdir = tempfile.mkdtemp()
    zin.extractall(tmpdir)

# Find the Dashboard sheet number (for sparklines)
# Sheet order: Raw_Data=1, Statistics=2, Pivot_Source=3, CAGR_Ranking=4,
#              Moving_Avg_MAE=5, Trend_Seasonal=6, Correlation=7, Dashboard=8
# We'll insert Pivot_Table as a new sheet

# Add pivot cache files
os.makedirs(f'{tmpdir}/xl/pivotCache', exist_ok=True)
os.makedirs(f'{tmpdir}/xl/pivotTables', exist_ok=True)

with open(f'{tmpdir}/xl/pivotCache/pivotCacheDefinition1.xml', 'w', encoding='utf-8') as f:
    f.write(cache_def)
with open(f'{tmpdir}/xl/pivotCache/pivotCacheRecords1.xml', 'w', encoding='utf-8') as f:
    f.write(records_xml)
with open(f'{tmpdir}/xl/pivotTables/pivotTable1.xml', 'w', encoding='utf-8') as f:
    f.write(pivot_table)

# Create pivot cache relationship
cache_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
    Target="pivotCacheRecords1.xml"/>
</Relationships>'''
os.makedirs(f'{tmpdir}/xl/pivotCache/_rels', exist_ok=True)
with open(f'{tmpdir}/xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', 'w') as f:
    f.write(cache_rels)

# Create pivot table relationship to cache
pt_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>'''
os.makedirs(f'{tmpdir}/xl/pivotTables/_rels', exist_ok=True)
with open(f'{tmpdir}/xl/pivotTables/_rels/pivotTable1.xml.rels', 'w') as f:
    f.write(pt_rels)

# Add pivot table to Pivot_Source sheet's relationships (sheet3.xml)
# Actually, let's add it to a dedicated sheet. But we already have 8 sheets.
# The pivot table should be on the Pivot_Source sheet or a separate Pivot_Table sheet.
# Let's add it as a relationship to sheet3 (Pivot_Source)

sheet3_rels_path = f'{tmpdir}/xl/worksheets/_rels/sheet3.xml.rels'
if os.path.exists(sheet3_rels_path):
    tree = etree.parse(sheet3_rels_path)
    root = tree.getroot()
else:
    root = etree.Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationships')
    tree = etree.ElementTree(root)
    os.makedirs(os.path.dirname(sheet3_rels_path), exist_ok=True)

# Add pivot table relationship
rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
existing_ids = [el.get('Id') for el in root]
new_id = f'rId{len(existing_ids)+1}'
new_rel = etree.SubElement(root, f'{{{rel_ns}}}Relationship')
new_rel.set('Id', new_id)
new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable')
new_rel.set('Target', '../pivotTables/pivotTable1.xml')
tree.write(sheet3_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# Update workbook.xml to add pivotCache
wb_path = f'{tmpdir}/xl/workbook.xml'
tree_wb = etree.parse(wb_path)
root_wb = tree_wb.getroot()
ns_main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Add pivotCaches element
pivot_caches = root_wb.find(f'{{{ns_main}}}pivotCaches')
if pivot_caches is None:
    # Insert after sheets element
    sheets_el = root_wb.find(f'{{{ns_main}}}sheets')
    idx = list(root_wb).index(sheets_el) + 1
    pivot_caches = etree.Element(f'{{{ns_main}}}pivotCaches')
    root_wb.insert(idx, pivot_caches)

cache_el = etree.SubElement(pivot_caches, f'{{{ns_main}}}pivotCache')
cache_el.set('cacheId', '0')
cache_el.set(f'{{{NS["r"]}}}id', 'rId_pivot')

tree_wb.write(wb_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# Add pivot cache relationship to workbook.xml.rels
wb_rels_path = f'{tmpdir}/xl/_rels/workbook.xml.rels'
tree_wbr = etree.parse(wb_rels_path)
root_wbr = tree_wbr.getroot()
rel_el = etree.SubElement(root_wbr, f'{{{rel_ns}}}Relationship')
rel_el.set('Id', 'rId_pivot')
rel_el.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition')
rel_el.set('Target', 'pivotCache/pivotCacheDefinition1.xml')
tree_wbr.write(wb_rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# Update [Content_Types].xml
ct_path = f'{tmpdir}/[Content_Types].xml'
tree_ct = etree.parse(ct_path)
root_ct = tree_ct.getroot()
ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'

overrides = [
    ('/xl/pivotCache/pivotCacheDefinition1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'),
    ('/xl/pivotCache/pivotCacheRecords1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'),
    ('/xl/pivotTables/pivotTable1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'),
]
for part, ctype in overrides:
    ov = etree.SubElement(root_ct, f'{{{ct_ns}}}Override')
    ov.set('PartName', part)
    ov.set('ContentType', ctype)

tree_ct.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)

print("✅ Pivot table XML injected")

# ══ SPARKLINES ══
# Add sparklines to Dashboard sheet (sheet8.xml)
# Each age group gets a sparkline showing year-over-year trend

sheet8_path = f'{tmpdir}/xl/worksheets/sheet8.xml'
tree_s8 = etree.parse(sheet8_path)
root_s8 = tree_s8.getroot()

# Sparkline data is in Dashboard sheet:
# Row SPARK_START+1 to SPARK_END, columns B to H (years), sparkline in column I+1
# Dashboard is the 8th sheet

spark_col = len(years) + 2  # column after years data (e.g. col 9 = I)
spark_col_letter = chr(64 + spark_col)  # 'I'
data_start_col = 'B'
data_end_col = chr(64 + len(years) + 1)  # 'H' for 7 years

sparkline_groups_xml = f'''<ext xmlns:x14="{NS['x14']}" uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}">
  <x14:sparklineGroups xmlns:xm="{NS['xm']}">
    <x14:sparklineGroup type="column" displayEmptyCellsAs="gap">
      <x14:colorSeries rgb="FF4472C4"/>
      <x14:colorNegative rgb="FFFF0000"/>
      <x14:colorAxis rgb="FF000000"/>
      <x14:sparklines>
'''

for i in range(len(age_groups)):
    data_row = SPARK_START + 1 + i
    sparkline_groups_xml += f'''        <x14:sparkline>
          <xm:f>Dashboard!{data_start_col}{data_row}:{data_end_col}{data_row}</xm:f>
          <xm:sqref>{spark_col_letter}{data_row}</xm:sqref>
        </x14:sparkline>
'''

sparkline_groups_xml += '''      </x14:sparklines>
    </x14:sparklineGroup>
  </x14:sparklineGroups>
</ext>'''

# Find or create extLst in sheet8
extLst = root_s8.find(f'{{{ns_main}}}extLst')
if extLst is None:
    extLst = etree.SubElement(root_s8, f'{{{ns_main}}}extLst')

# Parse sparkline XML and append
spark_el = etree.fromstring(sparkline_groups_xml)
extLst.append(spark_el)

tree_s8.write(sheet8_path, xml_declaration=True, encoding='UTF-8', standalone=True)
print("✅ Sparklines injected into Dashboard")

# ══ Repack ZIP ══
os.remove(TMP)
with zipfile.ZipFile(XLSX, 'w', zipfile.ZIP_DEFLATED) as zout:
    for root_dir, dirs, files in os.walk(tmpdir):
        for fname in files:
            fpath = os.path.join(root_dir, fname)
            arcname = os.path.relpath(fpath, tmpdir)
            zout.write(fpath, arcname)

# Cleanup
shutil.rmtree(tmpdir)
print(f"✅ Final xlsx saved: {XLSX}")

# Verify
with zipfile.ZipFile(XLSX, 'r') as z:
    names = z.namelist()
    pivot_files = [n for n in names if 'pivot' in n.lower()]
    spark_check = any('sparkline' in z.read(n).decode('utf-8', errors='ignore').lower()
                      for n in names if n.endswith('.xml') and 'sheet8' in n)
    print(f"   Pivot files: {pivot_files}")
    print(f"   Sparklines in sheet8: {spark_check}")
    print(f"   Total files: {len(names)}")
