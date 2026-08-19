import streamlit as st
import pandas as pd
import os
import tempfile

# =================================================================
# MOTOR DE PROCESSAMENTO (BACKEND)
# =================================================================
PROCESSOS_INTERNOS = [
    'calandra', 'corte-tubo', 'laser', 'laser / dobra', 
    'montagem', 'serralheria', 'solda', 'usinagem', 'usin. compl.'
]
PROCESSOS_EXTERNOS = [
    'eletroerosão', 'fundição', 'oxicorte', 'rebordear', 
    'revest. borracha', 'revest. escova', 'usin. polímero', 
    'usin. externa', 'water jet'
]
PROCESSOS_COMPRAS = [
    'comercial', 'itens de fixação', 'disp. eletromecânicos'
]

def is_numeric_sap(val):
    if pd.isna(val) or str(val).strip() in ['', '-', 'nan', '0']: return False
    val_str = str(val).split('.')[0].strip()
    return val_str.isnumeric()

def processar_planilha_pcp(caminho_arquivo, nome_montagem_final=None):
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    # Remove qualquer sufixo de arquivo temporário para manter o nome limpo
    nome_limpo = nome_base.split('_')[0] if '_' in nome_base else nome_base
    arquivo_saida = f"Estrutura {nome_limpo}.xlsx"

    df = pd.read_excel(caminho_arquivo)
    df.columns = df.columns.str.strip()
    df = df.fillna('')
    
    if 'Estrutura Solid' not in df.columns:
        raise ValueError("A coluna 'Estrutura Solid' não foi encontrada no arquivo.")
        
    df = df[df['Estrutura Solid'] != ''].copy()
    
    if nome_montagem_final and str(nome_montagem_final).strip() != "":
        df['Estrutura Solid'] = '    ' + df['Estrutura Solid'].astype(str)
        linha_mestre = pd.DataFrame({
            'Estrutura Solid': [str(nome_montagem_final).strip()], 
            'Quantidade': [1.0],
            'PROCESSO:': ['Montagem']
        })
        df = pd.concat([linha_mestre, df], ignore_index=True).fillna('')

    df['Item_Clean'] = df['Estrutura Solid'].astype(str).str.split('(').str[0].str.strip()
    df['indent'] = df['Estrutura Solid'].astype(str).apply(lambda x: len(x) - len(x.lstrip()))
    df['Level'] = df['indent'].map({indent: i for i, indent in enumerate(sorted(df['indent'].unique()))})
    df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(1.0)
    
    df['PROCESSO:'] = df.get('PROCESSO:', '').astype(str).str.strip().str.lower()
    df['MATÉRIA-PRIMA:'] = df.get('MATÉRIA-PRIMA:', '').astype(str).str.strip()
    df['CÓDIGO SAP:'] = df.get('CÓDIGO SAP:', '').apply(lambda x: str(x).split('.')[0] if is_numeric_sap(x) else '')

    df['Is_Comercial'] = df.apply(lambda row: True if any(p in row['PROCESSO:'] for p in PROCESSOS_COMPRAS) else False, axis=1)
    
    def tem_materia_prima_valida(mp):
        mp_clean = mp.split('(')[0].strip()
        return mp_clean not in ["", "-", "nan", "None"] and "COOL-" not in mp_clean and "DEP-" not in mp_clean and "FSK-" not in mp_clean

    df['Has_MP'] = df['MATÉRIA-PRIMA:'].apply(tem_materia_prima_valida)

    tem_filhos = [False] * len(df)
    lvls = df['Level'].tolist()
    is_comercial_list = df['Is_Comercial'].tolist()
    for i in range(len(lvls) - 1):
        if lvls[i+1] > lvls[i] and not is_comercial_list[i]: 
            tem_filhos[i] = True
    df['Tem_Filhos'] = tem_filhos

    def define_necessidade_op(row):
        if row['Is_Comercial']: return False
        if any(p in row['PROCESSO:'] for p in PROCESSOS_EXTERNOS): return False
        if row['Tem_Filhos']: return True
        if any(p in row['PROCESSO:'] for p in PROCESSOS_INTERNOS): return True
        return False
        
    df['Necessita_OP'] = df.apply(define_necessidade_op, axis=1)
    df['CÓDIGO_SAP_FINAL'] = df.apply(lambda row: row['CÓDIGO SAP:'] if row['Is_Comercial'] and row['CÓDIGO SAP:'] else row['Item_Clean'], axis=1)
    df = df.reset_index(drop=True)

    bom_lines = []
    seen_bom_parents = set()
    
    for i, row in df.iterrows():
        pai_sap = row['CÓDIGO_SAP_FINAL']
        if not row['Necessita_OP'] or pai_sap in seen_bom_parents: continue
            
        if row['Tem_Filhos']:
            filhos_diretos = []
            for j in range(i + 1, len(df)):
                child = df.iloc[j]
                if child['Level'] <= row['Level']: break
                if child['Level'] == row['Level'] + 1:
                    filhos_diretos.append(child)
            
            if filhos_diretos:
                seen_bom_parents.add(pai_sap)
                df_filhos = pd.DataFrame(filhos_diretos)
                parent_qty = row['Quantidade'] if row['Quantidade'] > 0 else 1.0
                df_filhos['Qtd_Base_Unitaria'] = df_filhos['Quantidade'] / parent_qty
                
                df_agrupados = df_filhos.groupby('CÓDIGO_SAP_FINAL').agg({'Item_Clean': 'first', 'Qtd_Base_Unitaria': 'sum'}).reset_index()
                for _, f_row in df_agrupados.iterrows():
                    bom_lines.append({
                        'Código Pai (OITT)': pai_sap,
                        'Código Componente (ITT1)': f_row['CÓDIGO_SAP_FINAL'],
                        'Descrição Componente': f_row['Item_Clean'],
                        'Quantidade Unitária': f_row['Qtd_Base_Unitaria']
                    })
                    
        elif row['Has_MP']:
            seen_bom_parents.add(pai_sap)
            mp_codigo_sap = row['CÓDIGO SAP:'] if is_numeric_sap(row['CÓDIGO SAP:']) else 'A DEFINIR'
            bom_lines.append({
                'Código Pai (OITT)': pai_sap,
                'Código Componente (ITT1)': mp_codigo_sap,
                'Descrição Componente': row['MATÉRIA-PRIMA:'],
                'Quantidade Unitária': 1.0
            })

    unique_oitm = []
    seen_oitm_codes = set()
    for _, row in df.iterrows():
        sap_code = row['CÓDIGO_SAP_FINAL']
        if not row['Is_Comercial'] and sap_code not in seen_oitm_codes:
            seen_oitm_codes.add(sap_code)
            tipo = 'DESENHO FABRICADO (TEM ESTRUTURA)' if row['Necessita_OP'] else 'DESENHO COMPRADO PRONTO (EXTERNO)'
            unique_oitm.append({
                'Código SAP (Desenho)': sap_code,
                'Descrição': row.get('Descrição', ''),
                'Tipo Cadastro': tipo,
                'Processo': row.get('PROCESSO:', '').upper(),
                'Tratamento Superficial': row.get('TRATAMENTO SUPERFICIAL:', ''),
                'Material Base': row.get('MATERIAL:_1', ''),
                'Matéria-Prima Descritiva': row.get('MATÉRIA-PRIMA:', ''),
                'Tamanho': row.get('Tamanho', ''),
                'Peso': row.get('Peso', '')
            })

    ops_df = df[df['Necessita_OP']].copy()
    if not ops_df.empty:
        df_ops_agregado = ops_df.groupby('CÓDIGO_SAP_FINAL').agg({
            'Item_Clean': 'first',
            'Quantidade': 'sum',
            'Level': 'max'
        }).reset_index()
        df_ops_agregado = df_ops_agregado.sort_values(by=['Level', 'Item_Clean'], ascending=[False, True]).reset_index(drop=True)
        df_ops_agregado.insert(0, 'Ordem Abertura', range(1, len(df_ops_agregado) + 1))
    else:
        df_ops_agregado = pd.DataFrame()

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        pd.DataFrame(unique_oitm).to_excel(writer, sheet_name='OITM_Cadastrar', index=False)
        pd.DataFrame(bom_lines).to_excel(writer, sheet_name='ITT1_Estrutura', index=False)
        if not df_ops_agregado.empty:
            df_ops_export = df_ops_agregado[['Ordem Abertura', 'CÓDIGO_SAP_FINAL', 'Item_Clean', 'Quantidade']].rename(
                columns={'Item_Clean': 'Item Limpo'}
            )
            df_ops_export.to_excel(writer, sheet_name='Plano_OPs', index=False)
            
    return arquivo_saida

# =================================================================
# INTERFACE DO APLICATIVO (FRONTEND)
# =================================================================
st.set_page_config(page_title="PCP | Integrador SAP", layout="centered", page_icon="⚙️")

st.title("⚙️ Gerar estrutura ➔ SAP")
st.markdown("Faça o upload da planilha com as informações do Solid.")

arquivo_upload = st.file_uploader("1. Selecione a planilha (.xlsx)", type=["xlsx"])
montagem_final = st.text_input("2. Código da Montagem Final (Opcional)", placeholder="Ex: FSK-1700")

if arquivo_upload is not None:
    if st.button("Processar Estrutura", type="primary"):
        with st.spinner("Limpando dados e aplicando regras de roteiro de produção..."):
            try:
                # O nome original do arquivo enviado pelo usuário
                nome_original = arquivo_upload.name 
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", prefix=nome_original.replace('.xlsx', '_')) as tmp:
                    tmp.write(arquivo_upload.getvalue())
                    tmp_path = tmp.name

                arquivo_saida = processar_planilha_pcp(tmp_path, montagem_final)

                with open(arquivo_saida, "rb") as file:
                    st.success(f"Tabelas geradas com sucesso! Arquivo: {arquivo_saida}")
                    st.download_button(
                        label="📥 Baixar Tabelas para o DTW",
                        data=file,
                        file_name=arquivo_saida,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                os.remove(tmp_path)
                os.remove(arquivo_saida)

            except Exception as e:
                st.error(f"Erro na conversão: {e}")