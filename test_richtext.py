import openpyxl
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.styles import colors

wb = openpyxl.Workbook()
ws = wb.active
red_font = InlineFont(color='FFFF0000')
black_font = InlineFont(color='FF000000')
rich_string = CellRichText(
    TextBlock(black_font, "PARCEL ID #"),
    TextBlock(red_font, "53-01031.000"),
    TextBlock(black_font, ": ")
)
ws['A1'] = rich_string
wb.save('/Users/davidmichalove/Desktop/automate/app/test_richtext.xlsx')
print("Done")
