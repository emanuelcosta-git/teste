import lxml.etree as ET
import openpyxl

# Caminho do arquivo XML e do Excel de saída
xml_path = "FC36182047000160_20250714_20250716155821_XZPROVFIM54654.xml"      # <-- Altere para o nome do seu arquivo XML
excel_path = "resultado.xlsx"    # <-- Nome do arquivo Excel de saída

# Parse do XML
tree = ET.parse(xml_path)
root = tree.getroot()

# Namespaces usados no XML
ns = {
    "ns": "http://www.sistemagalgo.com/SchemaPosicaoAtivos",
    "iso": "urn:iso:std:iso:20022:tech:xsd:semt.003.001.04"
}

# Extrai a data do extrato (vale para todos)
data_extrato = root.find('.//iso:StmtGnlDtls/iso:StmtDtTm/iso:Dt', ns)
data_extrato = data_extrato.text if data_extrato is not None else ""

# Encontra todos os BalForSubAcct
balfor_list = root.findall('.//iso:BalForSubAcct', ns)

dados = []
for balfor in balfor_list:
    # Fundo (Ticker)
    ticker = ""
    for othrid in balfor.findall('.//iso:FinInstrmId/iso:OthrId', ns):
        prtry = othrid.find('./iso:Tp/iso:Prtry', ns)
        if prtry is not None and prtry.text == "TICKER":
            ticker = othrid.find('./iso:Id', ns).text
            break

    # CNPJ
    cnpj = ""
    for othrid in balfor.findall('.//iso:FinInstrmId/iso:OthrId', ns):
        cd = othrid.find('./iso:Tp/iso:Cd', ns)
        if cd is not None and cd.text == "CNPJ":
            cnpj = othrid.find('./iso:Id', ns).text
            break

    # Patrimônio
    patrimonio = ""
    hldgval = balfor.find('.//iso:AcctBaseCcyAmts/iso:HldgVal/iso:Amt', ns)
    if hldgval is not None:
        patrimonio = hldgval.text

    # Quantidade de cotas
    qtd_cotas = ""
    unit = balfor.find('.//iso:AggtBal/iso:Qty/iso:Qty/iso:Qty/iso:Unit', ns)
    if unit is not None:
        qtd_cotas = unit.text

    dados.append([ticker, cnpj, patrimonio, data_extrato, qtd_cotas])

# Grava no Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "BalForSubAcct"
ws.append(["Fundo", "CNPJ", "Patrimonio", "Data", "Quantidade de cotas"])
for row in dados:
    ws.append(row)
wb.save(excel_path)

print(f"Extração concluída! Dados salvos em {excel_path}")