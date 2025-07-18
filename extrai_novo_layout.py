import openpyxl
from lxml import etree

xml_file = "FC36180712000187_20250710_20250711035033_XZFICFI54625.xml"
excel_file = "saida_novo_layout33.xlsx"

ns = {
    "sg": "http://www.sistemagalgo.com/SchemaPosicaoAtivos",
    "iso": "urn:iso:std:iso:20022:tech:xsd:semt.003.001.04"
}

tree = etree.parse(xml_file)

headers = [
    "FinInstrmId",
    "Ativo",
    "AggtBal.Qty.Unit (Quantidade)",
    "PricDtls.Val.Amt (Preço)",
    "AcctBaseCcyAmts.HldgVal.Amt (Valor Mercado)",
    "BalForSubAcct.FinInstrmId.Id (ID)",
    "SplmntryData.Sequencia (sq)",
    "StmtDtTm.Dt (Data Arquivo)",
    "Amt[@Ccy] (Moeda)"
]

wb = openpyxl.Workbook()
ws = wb.active
ws.append(headers)

data_arquivo = tree.find(".//iso:StmtGnlDtls/iso:StmtDtTm/iso:Dt", namespaces=ns)
data_arquivo = data_arquivo.text if data_arquivo is not None else ""

for sub in tree.xpath("//iso:BalForSubAcct", namespaces=ns):
    # ISIN
    isin = sub.find(".//iso:ISIN", namespaces=ns)
    isin = isin.text if isin is not None else ""

    # Ativo (Ticker)
    ticker = ""
    for othrid in sub.findall(".//iso:OthrId", namespaces=ns):
        prtry = othrid.find("./iso:Tp/iso:Prtry", namespaces=ns)
        if prtry is not None and prtry.text == "TICKER":
            ticker = othrid.find("./iso:Id", namespaces=ns).text
            break

    # Quantidade (Unit)
    quantidade = sub.find(".//iso:AggtBal//iso:Unit", namespaces=ns)
    quantidade = quantidade.text if quantidade is not None else ""

    # Preço (MRKT)
    preco = ""
    for prc in sub.findall(".//iso:PricDtls", namespaces=ns):
        tp = prc.find("./iso:Tp/iso:Cd", namespaces=ns)
        if tp is not None and tp.text == "MRKT":
            preco = prc.find("./iso:Val/iso:Amt", namespaces=ns).text
            break

    # Valor Mercado (HldgVal.Amt)
    valor_mercado = sub.find(".//iso:AcctBaseCcyAmts/iso:HldgVal/iso:Amt", namespaces=ns)
    valor_mercado = valor_mercado.text if valor_mercado is not None else ""

    # ID (primeiro OthrId/Id)
    id_ativo = ""
    othrid = sub.find(".//iso:OthrId/iso:Id", namespaces=ns)
    if othrid is not None:
        id_ativo = othrid.text

    # Sequencia (SplmntryData.Sequencia) - não existe no seu XML, ficará vazio
    sequencia = ""

    # Moeda (Ccy)
    moeda = ""
    moeda_elem = sub.find(".//iso:AcctBaseCcyAmts/iso:HldgVal/iso:Amt", namespaces=ns)
    if moeda_elem is not None and moeda_elem.get("Ccy"):
        moeda = moeda_elem.get("Ccy")
    else:
        moeda = "BRL"

    row = [
        isin,
        ticker,
        quantidade,
        preco,
        valor_mercado,
        id_ativo,
        sequencia,
        data_arquivo,
        moeda
    ]
    ws.append(row)

wb.save(excel_file)
print(f"Extração concluída! Dados salvos em {excel_file}") 