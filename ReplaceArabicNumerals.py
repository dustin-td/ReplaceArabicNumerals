import openpyxl
import sys, re

def ReplaceArabicNumerals(inbook,outbook,sheet):
    wb = openpyxl.load_workbook(inbook)
    sht = wb[sheet]

    arabicCells = SearchForArabic(sht)

    arabicNumerals = [(u'\u0660','0'),
                      (u'\u0661','1'),
                      (u'\u0662','2'),
                      (u'\u0663','3'),
                      (u'\u0664','4'),
                      (u'\u0665','5'),
                      (u'\u0666','6'),
                      (u'\u0667','7'),
                      (u'\u0668','8'),
                      (u'\u0669','9')]

    for cell in arabicCells:
        if cell.value is not None:
            val = unicode(cell.value)
            for k, v in arabicNumerals:
                val = val.replace(v, k)
            cell.value = val

    wb.save(outbook)

def SearchForArabic(worksheet):
    arabicCells = []
    pattern = re.compile(u'[\u0600-\u06FF]')
    for row in worksheet:
        for cell in row:
            val = unicode(cell.value)
            if pattern.search(val) is not None:
                arabicCells.append(cell)

    return arabicCells

if __name__ == '__main__':
    ReplaceArabicNumerals(sys.argv[1], sys.argv[2], sys.argv[3])