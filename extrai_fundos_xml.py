import xml.etree.ElementTree as ET
import pandas as pd

def get_text(element, path, ns):
    found = element.find(path, ns)
    return found.text if found is not None else ''

def formatar_monetario(valor):
    if not valor:
        return ''
    try:
        valor_float = float(valor)
        return f'{valor_float:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return valor

def formatar_cotas(valor):
    if not valor:
        return ''
    try:
        partes = valor.split('.')
        inteiro = partes[0]
        decimal = partes[1] if len(partes) > 1 else ''
        inteiro_formatado = f'{int(inteiro):,}'.replace(',', '.')
        if decimal:
            return f'{inteiro_formatado},{decimal}'
        else:
            return inteiro_formatado
    except Exception:
        return valor

def adicionar_total(df, colunas_soma, nome_coluna_total, cotas_cols=None):
    total = {}
    for col in df.columns:
        if col in colunas_soma:
            try:
                soma = pd.to_numeric(df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').sum()
                if cotas_cols and col in cotas_cols:
                    total[col] = formatar_cotas(f'{soma:.8f}'.rstrip('0').rstrip('.') if '.' in f'{soma:.8f}' else f'{soma:.8f}')
                else:
                    total[col] = formatar_monetario(soma)
            except Exception:
                total[col] = ''
        elif col == nome_coluna_total:
            total[col] = 'Total'
        else:
            total[col] = ''
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

def extrair_fundos_xml(xml_path):
    ns = {'ns': 'urn:iso:std:iso:20022:tech:xsd:semt.003.001.04'}
    tree = ET.parse(xml_path)
    root = tree.getroot()
    acoes = []
    fundos_imob = []
    fundos_rf = []
    fundos_rv = []
    for bal in root.findall('.//ns:BalForSubAcct', ns):
        isin = get_text(bal, './/ns:ISIN', ns)
        ticker = ''
        cnpj = ''
        for othrid in bal.findall('.//ns:OthrId', ns):
            tp = othrid.find('./ns:Tp/ns:Prtry', ns)
            if tp is not None and tp.text == 'TICKER':
                ticker = get_text(othrid, './ns:Id', ns)
            tp_cnpj = othrid.find('./ns:Tp/ns:Cd', ns)
            if tp_cnpj is not None and tp_cnpj.text == 'CNPJ':
                cnpj = get_text(othrid, './ns:Id', ns)
        nome = get_text(bal, './/ns:Desc', ns)
        qtd_disp = get_text(bal, './/ns:AggtBal//ns:Unit', ns)
        qtde_bloq = ''
        for brkdwn in bal.findall('.//ns:BalBrkdwn', ns):
            subbal = brkdwn.find('./ns:SubBalTp/ns:Cd', ns)
            if subbal is not None and subbal.text == 'BLOK':
                qtde_bloq = get_text(brkdwn, './/ns:Unit', ns)
        try:
            qtd_total = str(float(qtd_disp or 0) + float(qtde_bloq or 0))
        except Exception:
            qtd_total = qtd_disp or qtde_bloq
        valor_cota = ''
        for prc in bal.findall('.//ns:PricDtls', ns):
            tp = prc.find('./ns:Tp/ns:Cd', ns)
            if tp is not None and tp.text == 'MRKT':
                valor_cota = get_text(prc, './ns:Val/ns:Amt', ns)
        valor_atual = get_text(bal, './/ns:AcctBaseCcyAmts/ns:HldgVal/ns:Amt', ns)
        classificacao = get_text(bal, './/ns:ClssfctnFinInstrm', ns)
        is_acao = False
        for othrid in bal.findall('.//ns:OthrId', ns):
            tp = othrid.find('./ns:Tp/ns:Prtry', ns)
            if tp is not None and tp.text == 'ATIVOSB3':
                is_acao = True
        is_fii = ticker.endswith('11') if ticker else False
        is_rf = classificacao.startswith(('CDB', 'LF', 'LCI', 'LCA', 'DEB')) if classificacao else False
        if is_acao:
            acoes.append({
                'Cód.': ticker if ticker else isin,
                'Papel': nome,
                'Qtd. Disponível': qtd_disp,
                'Qtd. Bloqueada': qtde_bloq,
                'Qtd. Total': qtd_total,
                'Cotação': valor_cota,
                'Valor de Mercado Líquido': valor_atual
            })
        elif is_fii:
            fundos_imob.append({
                'Código': ticker if ticker else isin,
                'Fundo': nome,
                'Instituição': cnpj,
                'Quantidade Cotas': qtd_disp,
                'Qtde Bloq': qtde_bloq,
                'Valor Cota': valor_cota,
                'Valor Atual': valor_atual,
                'Valor Líquido': valor_atual
            })
        elif is_rf:
            fundos_rf.append({
                'Código': ticker if ticker else isin,
                'Fundo': nome,
                'Instituição': cnpj,
                'Quantidade Cotas': qtd_disp,
                'Qtde Bloq': qtde_bloq,
                'Valor Cota': valor_cota,
                'Valor Atual': valor_atual,
                'Valor Líquido': valor_atual
            })
        else:
            fundos_rv.append({
                'Código': ticker if ticker else isin,
                'Fundo': nome,
                'Instituição': cnpj,
                'Quantidade Cotas': qtd_disp,
                'Qtde Bloq': qtde_bloq,
                'Valor Cota': valor_cota,
                'Valor Atual': valor_atual,
                'Valor Líquido': valor_atual
            })
    df_acoes = pd.DataFrame(acoes)
    df_fii = pd.DataFrame(fundos_imob)
    df_rf = pd.DataFrame(fundos_rf)
    df_rv = pd.DataFrame(fundos_rv)
    # Formatar colunas
    for d, cotas_cols, monet_cols in [
        (df_acoes, ['Qtd. Disponível', 'Qtd. Bloqueada', 'Qtd. Total'], ['Cotação', 'Valor de Mercado Líquido']),
        (df_fii, ['Quantidade Cotas', 'Qtde Bloq'], ['Valor Cota', 'Valor Atual', 'Valor Líquido']),
        (df_rf, ['Quantidade Cotas', 'Qtde Bloq'], ['Valor Cota', 'Valor Atual', 'Valor Líquido']),
        (df_rv, ['Quantidade Cotas', 'Qtde Bloq'], ['Valor Cota', 'Valor Atual', 'Valor Líquido'])]:
        for col in cotas_cols:
            if col in d:
                d[col] = d[col].apply(formatar_cotas)
        for col in monet_cols:
            if col in d:
                d[col] = d[col].apply(formatar_monetario)
    # Adicionar totais
    if not df_acoes.empty:
        df_acoes = adicionar_total(df_acoes, ['Qtd. Disponível', 'Qtd. Bloqueada', 'Qtd. Total', 'Valor de Mercado Líquido'], 'Cód.', cotas_cols=['Qtd. Disponível', 'Qtd. Bloqueada', 'Qtd. Total'])
    if not df_fii.empty:
        df_fii = adicionar_total(df_fii, ['Quantidade Cotas', 'Qtde Bloq', 'Valor Atual', 'Valor Líquido'], 'Código', cotas_cols=['Quantidade Cotas', 'Qtde Bloq'])
    if not df_rf.empty:
        df_rf = adicionar_total(df_rf, ['Quantidade Cotas', 'Qtde Bloq', 'Valor Atual', 'Valor Líquido'], 'Código', cotas_cols=['Quantidade Cotas', 'Qtde Bloq'])
    if not df_rv.empty:
        df_rv = adicionar_total(df_rv, ['Quantidade Cotas', 'Qtde Bloq', 'Valor Atual', 'Valor Líquido'], 'Código', cotas_cols=['Quantidade Cotas', 'Qtde Bloq'])
    with pd.ExcelWriter('fundos_exportados.xlsx') as writer:
        if not df_acoes.empty:
            df_acoes.to_excel(writer, sheet_name='Renda Variável (AÇÕES)', index=False)
        if not df_fii.empty:
            df_fii.to_excel(writer, sheet_name='Fundos Imobiliários', index=False)
        if not df_rf.empty:
            df_rf.to_excel(writer, sheet_name='Fundos de Renda Fixa', index=False)
        if not df_rv.empty:
            df_rv.to_excel(writer, sheet_name='Fundos de Renda Variável', index=False)
    print('Exportado para fundos_exportados.xlsx')

def extrair_contas_pagar_receber(xml_path):
    ns = {'ns': 'urn:iso:std:iso:20022:tech:xsd:semt.003.001.04'}
    tree = ET.parse(xml_path)
    root = tree.getroot()
    lancamentos = []
    # Percorre todos os BalBrkdwn e AddtlBalBrkdwnDtls
    for balfor in root.findall('.//ns:BalForAcct', ns) + root.findall('.//ns:BalForSubAcct', ns):
        for balbrkdwn in balfor.findall('.//ns:BalBrkdwn', ns):
            # Lançamento principal
            prtry = balbrkdwn.find('./ns:SubBalTp/ns:Prtry', ns)
            cd = balbrkdwn.find('./ns:SubBalTp/ns:Cd', ns)
            if prtry is not None or cd is not None:
                # Descrição
                desc = ''
                if prtry is not None:
                    id_ = prtry.find('./ns:Id', ns)
                    sch = prtry.find('./ns:SchmeNm', ns)
                    desc = (id_.text if id_ is not None else '')
                    if sch is not None and sch.text:
                        desc += ' - ' + sch.text
                elif cd is not None:
                    desc = cd.text
                # Valor
                faceamt = balbrkdwn.find('.//ns:FaceAmt', ns)
                if faceamt is not None and faceamt.text:
                    valor = float(faceamt.text)
                    lancamentos.append({'Descrição Lançamento': desc, 'Valor Lançamento': valor})
            # Lançamentos adicionais
            for addtl in balbrkdwn.findall('./ns:AddtlBalBrkdwnDtls', ns):
                prtry = addtl.find('./ns:SubBalTp/ns:Prtry', ns)
                cd = addtl.find('./ns:SubBalTp/ns:Cd', ns)
                desc = ''
                if prtry is not None:
                    id_ = prtry.find('./ns:Id', ns)
                    sch = prtry.find('./ns:SchmeNm', ns)
                    desc = (id_.text if id_ is not None else '')
                    if sch is not None and sch.text:
                        desc += ' - ' + sch.text
                elif cd is not None:
                    desc = cd.text
                # Data
                data = addtl.find('./ns:SubBalAddtlDtls', ns)
                if data is not None and data.text:
                    desc += f' com pagamento {data.text}'
                # Valor
                faceamt = addtl.find('.//ns:FaceAmt', ns)
                if faceamt is not None and faceamt.text:
                    valor = float(faceamt.text)
                    lancamentos.append({'Descrição Lançamento': desc, 'Valor Lançamento': valor})
    if not lancamentos:
        return None
    df = pd.DataFrame(lancamentos)
    # Total para %
    total = df['Valor Lançamento'].sum()
    df['% S/ CPR'] = df['Valor Lançamento'] / total * 100
    df['% S/ CPR'] = df['% S/ CPR'].map(lambda x: f'{x:.2f}%')
    # % S/ Total (usar o mesmo total, pois não há outro contexto)
    df['% S/ Total'] = df['Valor Lançamento'] / total * 100
    df['% S/ Total'] = df['% S/ Total'].map(lambda x: f'{x:.2f}%')
    # Formatar valores
    df['Valor Lançamento'] = df['Valor Lançamento'].map(lambda x: f'{x:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))
    # Adicionar total
    total_row = {'Descrição Lançamento': 'Total',
                 'Valor Lançamento': f'{total:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'),
                 '% S/ CPR': '100.00%', '% S/ Total': '100.00%'}
    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df

if __name__ == '__main__':
    extrair_fundos_xml('FC36182047000160_20250714_20250716155821_XZPROVFIM54654.xml')
    df_contas = extrair_contas_pagar_receber('FC36182047000160_20250714_20250716155821_XZPROVFIM54654.xml')
    with pd.ExcelWriter('fundos_exportados.xlsx', mode='a', if_sheet_exists='replace') as writer:
        if df_contas is not None:
            df_contas.to_excel(writer, sheet_name='Contas a Pagar_Receber', index=False) 