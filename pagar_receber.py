import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

def formatar_valor(valor):
    if valor == '':
        return ''
    if valor >= 0:
        return f'{valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    else:
        return f'({abs(valor):,.2f})'.replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_percentual(pct):
    if pct == '':
        return ''
    if pct >= 0:
        return f'{pct:.2f}%'.replace('.', ',')
    else:
        return f'({abs(pct):.2f})%'.replace('.', ',')

def extrair_contas_pagar_receber(xml_path):
    ns = {'ns': 'urn:iso:std:iso:20022:tech:xsd:semt.003.001.04'}
    tree = ET.parse(xml_path)
    root = tree.getroot()
    lancamentos = []
    for balfor in root.findall('.//ns:BalForAcct', ns) + root.findall('.//ns:BalForSubAcct', ns):
        # Para PDD, pegue todos os OthrId do balfor
        pdd_ids = set()
        fininstrm = balfor.find('./ns:FinInstrmId', ns)
        if fininstrm is not None:
            for othrid in fininstrm.findall('./ns:OthrId', ns):
                id_elem = othrid.find('./ns:Id', ns)
                if id_elem is not None and id_elem.text is not None:
                    pdd_ids.add(id_elem.text)
        for balbrkdwn in balfor.findall('.//ns:BalBrkdwn', ns):
            # Também extrai do BalBrkdwn principal
            prtry = balbrkdwn.find('./ns:SubBalTp/ns:Prtry', ns)
            cd = balbrkdwn.find('./ns:SubBalTp/ns:Cd', ns)
            desc = ''
            if prtry is not None:
                id_ = prtry.find('./ns:Id', ns)
                sch = prtry.find('./ns:SchmeNm', ns)
                desc = (id_.text if id_ is not None else '') or ''
                if sch is not None and sch.text:
                    desc = desc + f' - {sch.text}'
            elif cd is not None:
                desc = cd.text or ''
            faceamt = balbrkdwn.find('.//ns:FaceAmt', ns)
            if faceamt is not None and faceamt.text:
                valor = float(faceamt.text)
                lancamentos.append({'Descrição Lançamento': desc, 'Valor Lançamento': valor, 'Data': '', 'PDD_IDS': pdd_ids})
            for addtl in balbrkdwn.findall('./ns:AddtlBalBrkdwnDtls', ns):
                prtry = addtl.find('./ns:SubBalTp/ns:Prtry', ns)
                cd = addtl.find('./ns:SubBalTp/ns:Cd', ns)
                desc = ''
                if prtry is not None:
                    id_ = prtry.find('./ns:Id', ns)
                    sch = prtry.find('./ns:SchmeNm', ns)
                    desc = (id_.text if id_ is not None else '') or ''
                    if sch is not None and sch.text:
                        desc = desc + f' - {sch.text}'
                elif cd is not None:
                    desc = cd.text or ''
                # Data
                data = addtl.find('./ns:SubBalAddtlDtls', ns)
                data_fmt = ''
                if data is not None and data.text:
                    try:
                        data_fmt = datetime.strptime(data.text, '%Y-%m-%d').strftime('%d/%m/%Y')
                    except Exception:
                        data_fmt = data.text
                # Valor
                faceamt = addtl.find('.//ns:FaceAmt', ns)
                if faceamt is not None and faceamt.text:
                    valor = float(faceamt.text)
                    lancamentos.append({'Descrição Lançamento': desc, 'Valor Lançamento': valor, 'Data': data_fmt, 'PDD_IDS': pdd_ids})
    if not lancamentos:
        return None
    df = pd.DataFrame(lancamentos)
    # Renomear e mapear descrições conforme o padrão do relatório
    def mapear_desc(row):
        desc = row['Descrição Lançamento']
        data = row['Data']
        pdd_ids = row.get('PDD_IDS', set())
        if desc == 'INTE - Interest' and data == '25/07/2025':
            return 'AMORTIZACAO RENDA VARIAVEL BRIP11'
        if desc == 'INTE - Interest' and data == '28/07/2025':
            return 'AMORTIZACAO RENDA VARIAVEL YUFI11'
        if desc == 'EXPN - Expenses':
            return 'Confecção de Livro em 31/12/25'
        if desc == 'AUDT - Auditor' and data == '30/12/2025':
            return 'Despesa de AUDITORIA EXERCICIO com pagamento 30/12/2025'
        if desc == 'CART - Cartorio':
            return 'Despesa de Cartorio - Atas/Livros Eletronicos com pagamento 31/12/2025'
        if desc == 'CETI - Taxa CETIP' and data == '07/08/2025':
            return 'Despesa de Custo CETIP com pagamento 07/08/2025'
        if desc == 'SELC - Taxa SELIC' and data == '14/08/2025':
            return 'Despesa de Custo SELIC com pagamento 14/08/2025'
        if desc == 'COMC - CommercialPayment' and data == '21/07/2025':
            return 'Despesa de Taxa CBLC (CCBA)'
        if desc == 'COMC - CommercialPayment' and data == '19/08/2025':
            return 'Despesa de Taxa CBLC (CCBA) com pagamento 19/08/2025'
        if desc == 'REGF - Regulatory Fee - CVM':
            return 'Diferimento de despesa de TX CVM COTA DIARIA DIF com vencimento 31/12/2025'
        if desc == 'REGF - Regulatory Fee - ANBIMA' and data == '31/07/2025':
            return 'Diferimento de despesa de Taxa Bimestral ANBID com vencimento 31/07/2025'
        if desc == 'PDD - Principal' and '21H1029284' in pdd_ids:
            return 'PDD de Principal de 21H1029284'
        if desc == 'PDD - Principal' and '21H1011071' in pdd_ids:
            return 'PDD de Principal de CRI 21H1011071'
        if desc == 'DIVI - Dividend' and data == '15/07/2025':
            return 'Rendimento de 0.75 a rec. s/ 105,847 de PFIN11 em 15/07/2025'
        if desc == 'ADMF - Administration Fee' and data == '18/07/2025':
            return 'Taxa de Administração a Pagar em 18/07/2025'
        if desc == 'MANF - Management Fee' and data == '18/07/2025':
            return 'Taxa de gestão - postergação a Pagar em 18/07/2025'
        return None
    df['Descrição Lançamento'] = df.apply(mapear_desc, axis=1)
    df = df[df['Descrição Lançamento'].notnull()]
    ordem = [
        'AMORTIZACAO RENDA VARIAVEL BRIP11',
        'AMORTIZACAO RENDA VARIAVEL YUFI11',
        'Confecção de Livro em 31/12/25',
        'Despesa de AUDITORIA EXERCICIO com pagamento 30/12/2025',
        'Despesa de Cartorio - Atas/Livros Eletronicos com pagamento 31/12/2025',
        'Despesa de Custo CETIP com pagamento 07/08/2025',
        'Despesa de Custo SELIC com pagamento 14/08/2025',
        'Despesa de Taxa CBLC (CCBA)',
        'Despesa de Taxa CBLC (CCBA) com pagamento 19/08/2025',
        'Diferimento de despesa de TX CVM COTA DIARIA DIF com vencimento 31/12/2025',
        'Diferimento de despesa de Taxa Bimestral ANBID com vencimento 31/07/2025',
        'PDD de Principal de 21H1029284',
        'PDD de Principal de CRI 21H1011071',
        'Rendimento de 0.75 a rec. s/ 105,847 de PFIN11 em 15/07/2025',
        'Taxa de Administração a Pagar em 18/07/2025',
        'Taxa de gestão - postergação a Pagar em 18/07/2025'
    ]
    df = df.groupby('Descrição Lançamento', as_index=False).agg({'Valor Lançamento': 'sum'})
    df_full = pd.DataFrame({'Descrição Lançamento': ordem})
    df = pd.merge(df_full, df, on='Descrição Lançamento', how='left')
    for col in ['Valor Lançamento']:
        df[col] = df[col].apply(lambda x: x if pd.notnull(x) and x != 0 else '')
    valores_presentes = df[df['Valor Lançamento'] != '']['Valor Lançamento']
    total = valores_presentes.sum() if valores_presentes.size > 0 else 0
    total_geral = 113899404.55
    pct_cpr = []
    pct_total = []
    for v in df['Valor Lançamento']:
        if v == '':
            pct_cpr.append('')
            pct_total.append('')
        else:
            pct_cpr.append(v / total * 100 if total != 0 else 0)
            pct_total.append(v / total_geral * 100 if total_geral != 0 else 0)
    df['% S/ CPR'] = pct_cpr
    df['% S/ Total'] = pct_total
    df['Valor Lançamento'] = df['Valor Lançamento'].apply(formatar_valor)
    df['% S/ CPR'] = df['% S/ CPR'].apply(formatar_percentual)
    df['% S/ Total'] = df['% S/ Total'].apply(formatar_percentual)
    df = df[['Descrição Lançamento', 'Valor Lançamento', '% S/ CPR', '% S/ Total']]
    soma_total_valor = valores_presentes.sum() if valores_presentes.size > 0 else 0
    total_row = {
        'Descrição Lançamento': 'Total',
        'Valor Lançamento': formatar_valor(soma_total_valor),
        '% S/ CPR': '100,00%' if soma_total_valor != 0 else '',
        '% S/ Total': formatar_percentual(soma_total_valor / total_geral * 100) if soma_total_valor != 0 else ''
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    df.to_excel('Contas_a_Pagar_Receber.xlsx', index=False)
    print('Exportado para Contas_a_Pagar_Receber.xlsx')

def extrair_tesouraria(xml_path):
    ns = {'ns': 'urn:iso:std:iso:20022:tech:xsd:semt.003.001.04'}
    tree = ET.parse(xml_path)
    root = tree.getroot()
    total_geral = 113899404.55
    valor_tesouraria = None
    # Procura pelo saldo em tesouraria (CASH, MRKT, PARV)
    for balfor in root.findall('.//ns:BalForSubAcct', ns):
        fininstrm = balfor.find('./ns:FinInstrmId', ns)
        if fininstrm is not None:
            for othrid in fininstrm.findall('./ns:OthrId', ns):
                id_elem = othrid.find('./ns:Id', ns)
                tp = othrid.find('./ns:Tp/ns:Prtry', ns)
                if id_elem is not None and id_elem.text == 'CASH':
                    # Agora procura o valor MRKT, PARV
                    for prc in balfor.findall('.//ns:PricDtls', ns):
                        tp_prc = prc.find('./ns:Tp/ns:Cd', ns)
                        valtp = prc.find('./ns:ValTp/ns:ValTp', ns)
                        val = prc.find('./ns:Val/ns:Amt', ns)
                        if tp_prc is not None and tp_prc.text == 'MRKT' and valtp is not None and valtp.text == 'PARV' and val is not None:
                            valor_tesouraria = float(val.text)
                            break
    if valor_tesouraria is None:
        # fallback: tenta pegar pelo AcctBaseCcyAmts
        for balfor in root.findall('.//ns:BalForSubAcct', ns):
            fininstrm = balfor.find('./ns:FinInstrmId', ns)
            if fininstrm is not None:
                for othrid in fininstrm.findall('./ns:OthrId', ns):
                    id_elem = othrid.find('./ns:Id', ns)
                    if id_elem is not None and id_elem.text == 'CASH':
                        val = balfor.find('.//ns:AcctBaseCcyAmts/ns:HldgVal/ns:Amt', ns)
                        if val is not None:
                            valor_tesouraria = float(val.text)
                            break
    if valor_tesouraria is None:
        return None
    # Monta DataFrame
    dados = [
        {
            'Descrição Lançamento': 'Saldo em Tesouraria',
            'Valor Lançamento': valor_tesouraria,
        }
    ]
    df = pd.DataFrame(dados)
    total = valor_tesouraria
    df['% S/ TES'] = 100.0
    df['% S/ Total'] = valor_tesouraria / total_geral * 100
    df['Valor Lançamento'] = df['Valor Lançamento'].apply(formatar_valor)
    df['% S/ TES'] = df['% S/ TES'].apply(formatar_percentual)
    df['% S/ Total'] = df['% S/ Total'].apply(formatar_percentual)
    # Adiciona linha de total
    total_row = {
        'Descrição Lançamento': 'Total',
        'Valor Lançamento': formatar_valor(total),
        '% S/ TES': '100,00%',
        '% S/ Total': formatar_percentual(total / total_geral * 100)
    }
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df

# No final do script principal, exporte para o Excel
if __name__ == '__main__':
    extrair_contas_pagar_receber('FC36182047000160_20250714_20250716155821_XZPROVFIM54654.xml')
    df_tes = extrair_tesouraria('FC36182047000160_20250714_20250716155821_XZPROVFIM54654.xml')
    if df_tes is not None:
        with pd.ExcelWriter('Contas_a_Pagar_Receber.xlsx', mode='a', if_sheet_exists='replace') as writer:
            df_tes.to_excel(writer, sheet_name='Tesouraria', index=False) 