similar = lambda a, b: similar(a.value, b.value)

def get_ticker(sheet, coordinate, column):
    return str(sheet[column + coordinate[1:]].value)

matching_pairs = [(trad, green) for trad in tradnames for green in greenNames if similar(trad, green) > 0.9]

for trad, green in matching_pairs:
    tradric = get_ticker(tradsheet, trad.coordinate, 'K')
    greenric = get_ticker(greenSheet, green.coordinate, 'D')
    try:
        tradDataframe, error = rdp.get_data(
            instruments=tradric,
            fields=fields,
        )
        greenDataFrame, error = rdp.get_data(
            instruments=greenric,
            fields=fields,
        )
        sheet_name = str(greenSheet['k' + green.coordinate[1:]].value)
        tradDataframe.to_excel(writer, sheet_name=sheet_name, startcol=0)
        greenDataFrame.to_excel(writer, sheet_name=sheet_name, startcol=10)
        writer.save()
    except Exception as e:
        print(e)
        pass