
import pandas as pd



def createT1(file):
    # read xls ,save specific columns.
    with pd.ExcelFile(file) as xls:
        packList = pd.read_excel(xls, 'Packing List')
        civ = pd.read_excel(xls, 'Commercial Invoice')
    # print(packList)

    startIndexPacklist = 7
    startIndexCiv = 8

    civLen, packLen = getLengthOfColumns(packList, civ)
    Itemlist = packList[startIndexPacklist:packLen]['Unnamed: 2'].reset_index(drop=True)  # ready
    CTN = packList[startIndexPacklist:packLen]['Unnamed: 6'].reset_index(drop=True)  # ready
    GrossWeight = packList[startIndexPacklist:packLen]['Unnamed: 9'].reset_index(drop=True)  # ready
    NetWeight = packList[startIndexPacklist:packLen]['Unnamed: 11'].reset_index(drop=True)  # ready
    CustomsValue3 = civ[startIndexCiv:7 + civLen]['Unnamed: 6'].reset_index(drop=True)  # ready
    hsCode = civ[startIndexCiv:7 + civLen]['Unnamed: 9'].reset_index(drop=True)  # ready
    # Erstellung der Tabelle

    dict = {'H.S code': hsCode, 'Warenbeschreibung': Itemlist, 'Packstückanzahl': CTN, 'Gewicht': GrossWeight,
            'Netto': NetWeight, 'Warenwert': CustomsValue3}
    df = pd.DataFrame(dict)

    # Erstellung der T1
    T1 = df.reset_index(drop=True)
    print(T1)

    # Erstellung der Verzollung
    verzollung = createVerzollung(T1)
    print(type(verzollung))
    #createWorkbook(T1, verzollung, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    return T1, verzollung


def writeToExcel(writer,verzollung, T1):#DONT TOUCH
    verzollung.to_excel(writer, sheet_name='verzollung', startrow=9, startcol=0, index=True)#Index True weil HScode
    T1.to_excel(writer, sheet_name='T1', startrow=9, startcol=0, index=False)

    return



def createVerzollung(T1):  # DONT TOUCH
    join_unique = lambda x: ','.join(set(x))
    test = T1.groupby('H.S code').agg({'Warenbeschreibung': join_unique})  # erzeugt Warenbeschreibung für jede HS
    verzollung = T1.groupby(['H.S code']).sum()  # Summiert alles für jede HS
    verzollung['Warenbeschreibung'] = test['Warenbeschreibung']  # ersetzt Warenbeschreibung mit richtiger.
    # print(test)
    # print(verzollung)
    return verzollung


def getLengthOfColumns(packList, civ):  # DONT TOUCH
    lengthCiv = len(civ['Unnamed: 9'].dropna())
    lengthPackList = len(packList['Unnamed: 6']) - 1
    # print('lengthPacklist:',lengthPackList)
    return lengthCiv, lengthPackList


def createWorkbook(T1, verzollung, Sendungsnr, Schiff, BL,BLDatum, Rechnungsnr, RechnungsDatum, Containernr, Incoterm,
                   Transportpreis, Inlandpreis,writer):

    writeToExcel(writer, verzollung, T1)
    print('verzollung', verzollung)


    T1Sheet = writer.sheets['T1']
    VerzollungsSheet = writer.sheets['verzollung']
    T1Sheet.set_column(0, 10, 25)
    VerzollungsSheet.set_column(0, 10, 25)

    T1Sheet.write('A1', 'Sendungsnummer:')
    T1Sheet.write('B1', Sendungsnr)

    T1Sheet.write('A2', 'Schiff:')
    T1Sheet.write('B2', Schiff)

    T1Sheet.write('A3', 'B/L Nummer:')
    T1Sheet.write('B3', BL)
    T1Sheet.write('C3', BLDatum)

    T1Sheet.write('A4', 'Rechnung:')
    T1Sheet.write('B4', Rechnungsnr)
    T1Sheet.write('C4', RechnungsDatum)

    T1Sheet.write('A5', 'Container:')
    T1Sheet.write('B5', Containernr)

    T1Sheet.write('A6', 'Incoterm:')
    T1Sheet.write('B6', Incoterm)

    T1Sheet.write('A7', 'Transportpreis:')
    T1Sheet.write('B7', Transportpreis)

    T1Sheet.write('A8', 'Inland:')
    T1Sheet.write('B8', Inlandpreis)


    VerzollungsSheet.write('A1', 'Sendungsnummer:')
    VerzollungsSheet.write('B1', Sendungsnr)

    VerzollungsSheet.write('A2', 'Schiff:')
    VerzollungsSheet.write('B2', Schiff)

    VerzollungsSheet.write('A3', 'B/L Nummer:')
    VerzollungsSheet.write('B3', BL)
    VerzollungsSheet.write('C3', BLDatum)

    VerzollungsSheet.write('A4', 'Rechnung:')
    VerzollungsSheet.write('B4', Rechnungsnr)
    VerzollungsSheet.write('B4', RechnungsDatum)

    VerzollungsSheet.write('A5', 'Container:')
    VerzollungsSheet.write('B5', Containernr)
    VerzollungsSheet.write('A6', 'Incoterm:')
    VerzollungsSheet.write('B6', Incoterm)

    VerzollungsSheet.write('A7', 'Transportpreis:')
    VerzollungsSheet.write('B7', Transportpreis)

    VerzollungsSheet.write('A8', 'Inland:')
    VerzollungsSheet.write('B8', Inlandpreis)
    writer.close()
    return


#DAS
#Sendungsnr = str(input())
#Schiff = str(input())
#BL = str(input())
#BLDatum = str(input())
#Rechnungsnr = str(input())
#Rechnungsdatum = str(input())
#Containernr = str(input())
#Incoterm = "FOB "+str(input())
#Transportpreis = str(input())+" EUR"
#Inlandpreis = str(input())+" EUR"

#writer = pd.ExcelWriter(Sendungsnr+".xlsx", engine='xlsxwriter')
#T1, verzollung = createT1('test.xlsx')
#createWorkbook(T1, verzollung, Sendungsnr, Schiff, BL, BLDatum ,Rechnungsnr,Rechnungsdatum,Containernr, Incoterm, Transportpreis, Inlandpreis)
