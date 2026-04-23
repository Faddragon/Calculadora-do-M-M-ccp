import pandas as pd
import os
from rapidfuzz import process, fuzz

# =====================================================
# ARQUIVOS
# =====================================================
ARQUIVOS_PASTA = r"F:\documentos francisco\Trabalho\DataScience2\Morbimortalidade CCP IAVC\cp - organizador planilhas para MM"
ARQ_MEDICOS = os.path.join(ARQUIVOS_PASTA, 'medicos.xlsx')
ARQ_SAIDA = os.path.join(ARQUIVOS_PASTA, 'cirurgias_cp_MM.xlsx')
ARQ_MAPEAamentos = os.path.join(ARQUIVOS_PASTA, 'mapeamentos.xlsx')

# =====================================================
# CARREGAR MÉDICOS VÁLIDOS (da planilha ou lista padrão)
# =====================================================
def carregar_medicos():
    if os.path.exists(ARQ_MEDICOS):
        df = pd.read_excel(ARQ_MEDICOS)
        # Combinar NOME e CRM
        completa = df['NOME'] + ' (CRM - ' + df['CRM'].astype(str) + ')'
        return completa.tolist()
    else:
        return [
            "FRANCISCO ARAUJO DIAS (CRM - 154295)",
            "MARCELO SOARES SCHALCH (CRM - 164050)",
            "RAFAEL DE CICCO (CRM - 112733)",
            "VICTOR BANDINI VIEIRA (CRM - 164000)",
            "PABLO GABRIEL OCAMPO QUINTANA (CRM - 138741)",
            "ERICA ERINA FUKUYAMA (CRM - 72507)",
            "SANDRA CARINA LOPEZ CALCINES (CRM - 210541)",
            "GIOVANNA MARCELA VIEIRA DELLA NEGRA (CRM - 207914)",
        ]

def salvar_medico(nome, crm):
    """Adiciona novo médico à planilha."""
    novo = pd.DataFrame({'NOME': [nome], 'CRM': [crm]})
    
    if os.path.exists(ARQ_MEDICOS):
        df_existente = pd.read_excel(ARQ_MEDICOS)
        df_novo = pd.concat([df_existente, novo], ignore_index=True)
    else:
        df_novo = novo
    
    df_novo.to_excel(ARQ_MEDICOS, index=False)
    print(f"✅ Médico cadastrado: {nome} (CRM - {crm})")
    return carregar_medicos()

MEDICOS_VALIDOS = carregar_medicos()

# =====================================================
# AGRUPAMENTOS DE CIRURGIAS (dicionário completo com variações)
# =====================================================
agrupamentos = {
    'TIREOIDECTOMIA TOTAL': [
        'TIREOIDECTOMIA TOTAL', 'TIREOIDECTOMIA TOTAL + LINFADENECTOMIA RECORRENCIAL',
        'TIREOIDECTOMIA TOTAL (TOTALIZAÇÃO) + LINFADENECTOMIA RECORRENCIAL DIR.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO NIVEL VI DIR.',
        'TIREOIDECTOMIA TOTAL + EC NIVEL VI ESQ.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO RECORRENCIAL DIR. + DRENAGEM COM PORTOVAC',
        'TIREIDECTOMIA TOTAL+ LINFADENECTOMIA RECORRENCIAL',
        'TIRERECTOMIA TOTAL EM ONCOLOGIA +LINFANDENECTOMIA CERVICAL RECORRENCIAL',
        'TIREOIDECTOMIA TOTAL + EC NIVEL VI DIREITO',
        'TIREOIDECTOMIA TOTAL + LINFADENECTOMIA RECORRENCIAL (AMOSTRAGEM) + DRENAGEM PORTOVAC BOCIO MULTINODULAR',
        'TIREOIDECTOMIA TOTAL + EC NIVEL VI',
    ],
    'TIREOIDECTOMIA PARCIAL': [
        'TIREOIDECTOMIA PARCIAL', 'TIREOIDECTOMIA PARCIAL DIR. + ISTMO',
        'TIREOIDECTOMIA PARCIAL DIR. + LINFADENECTOMIA NIVEL VI',
        'TIREOIDECTOMIA PARCIAL ESQ. + ISTMO', 'TIREOIDECTOMIA PARCAIL ESQ.',
        'TIREOIDECTOMIA PARCIAL ESQ. + LINFADENECTOMIA',
        'TIREOIDECTOMIA PARCIAL - VLP',
        'TIREOIDECTOMIA PARCIAL ESQ. + LINFADENECTOMIA RECORRENCIAL',
        'TIREOIDECTOMIA PARCIAL DIR. + ISTMO + LINFADENECTOMIA RECORRENCIAL',
    ],
    'ABLAÇÃO DE NODULO TIREOIDIANO': [
        'ABLAÇÃO DE NODULO TIREOIDIANO POR RADIOFREQUENCIA',
        'ABLAÇÃO DE NODULO POR RADIO FREQUENCIA',
        'TIREOIDECTOMIA PACIAL+ABLA DE NODULO POR MICROONDAS',
        'TIREOIDECTOMIA PARCIAL DIREITA +ABLAÇÃO DENODULO TIREOIDIANO POR RADIOFREQUENCIA',
        'TIREOIDECTOMIA PARCIAL ESQUERDA COM ABLAÇÃO DE NODULO POR MICROONDAS',
        'ABLACAO DE NODULO EM LOBO ESQ. (MICROWAVE)',
        'TIREOIDECTOMIA PARCIAL - RADIOFREQUÊNCIA', 'TIREOIDECTOMIA PARCIAL - RADIOFREQUENCIA',
        'ABLACAO DE NODULO POR RADIOFREQUENCIA', 'ABLACAO DE NODULOS POR MICROONDAS',
        'ABLACAO DE NODULOS POR RADIOFREQUENCIA', 'ABLACAO DE NODULO POR MICROONDAS',
        'TIREOIDECTOMIA PARCIAL (RF)', 'AMBLACAO DE NODULO POR MICROWAVE',
    ],
    'TIREOIDECTOMIA TOTAL + EC LATERAL': [
        'TIREOIDECTOMIA TOTAL + EC NIVEIS II A IV + VI + TQT + DRENAGEM COM PORTOVAC',
        'TIREOIDECTIA TOTAL EM ONC+ LINFADENECTOMIA RADICAL MODIFICA',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL NIVEIS II, III, IV, V E VI A ESQ.',
        'TIREOIDECTOMIA TOTAL + LINFADENECTOMIA RADICAL MODIFICADA BILATERAL',
        'TIREOIDECTOMIA TOTAL + EC NIVEIS II, III, IV E VI DIR.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL II-IV DIR.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL II - IV BILATERAL',
        'TIREOIDECTOMIA TOTAL + EC NIVEL VI DIREITO',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL NIVEIS II-VI ESQ.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL NIVEIS II, III, IV, V E VI DIREITO',
        'TIREOIDECTOMIA TOTAL + EC II-IV BILATERAL + NIVEL VI BILATERAL',
        'TIREOIDECTOMIA TOTAL + EC II - IV ESQ.',
        'TIREOIDECTOMIA TOTAL + ESVAZIAMENTO CERVICAL NIVEIS II, III, IV, V E VI A ESQ. + TQT + SNE',
        'TIREOIDECTOMIA PARCIAL ESQ. + ESVAZIAMENTO CERVICAL NIVEIS II A VI A ESQ. NIVEL VI A DIR',
    ],
    'GLOSSECTOMIA PARCIAL': [
        'GLOSSECTOMIA PARCIAL', 'GLOSSECTOMIA PARCIAL ANTERIOR',
        'GLOSSECTOMIA PARCIAL ESQ.',
        'GLOSSECTOMIA PARCIAL ESQ. + ESVAZIAMENTO CERVICAL SUPRA-OMOHIOIDEO ESQ.',
        'GLOSSECTOMIA SUBTOTAL + ECI A IV A DIREITA',
        'GLOSSECTOMIA PARCIAL A DIR. + EC I A IV A ESQ.',
        'GLOSSECTOMIA SUBTOTAL + I A IV DIR. + EC I - III ESQ.',
    ],
    'GLOSSECTOMIA TOTAL': [
        'GLOSSECTOMIA TOTAL', 'GLOSSECTOMIA TOTAL + LINFADENECTOMIA RADICAL MODIFICADA ESQ.',
    ],
    'PGM': [
        'PELGLOSSOMANDIBULECTOMIA', 'PELVIGLOSSOMANDIBULECTOMIA',
        'PELVIGLOSSOMANDIBULECTOMIA ESQ.',
        'PGM + RECONSTRUCAO C/ RETALHO MIOCUTANEO',
        'PGM ANTERIOR + GLOSSECTOMIA TOTAL',
        'PGM + ECSOH BILATERAL',
        'PELVEGLOSSOMANDIBULECTOMIA ANTERIOR + MANDIBULECTOMIA MARGINAL',
    ],
    'MAXILECTOMIA': ['MAXILECTOMIA PARCIAL', 'RESSECCAO DE LESAO EM PALATO DURO'],
    'TRAQUEOSTOMIA': ['TRAQUEOSTOMIA', 'TRAQUEOSTOMIA DE URGENCIA', 'LARINGECTOMIA PARCIAL + TRAQUEOSTOMIA'],
    'PROVOX': ['PASSAGEM DE PROTESE PROVOX', 'TROCA DE PROVOX', 'RECONSTRUCAO PARA FONACAO', 'RECONTRUCAO PARA FONACAO', 'RECONSTRUÇÃO PARA FONAÇÃO'],
    'PAROTIDECTOMIA': [
        'PAROTIDECTOMIA', 'PAROTIDECTOMIA TOTAL', 'PAROTIDECTOMIA TOTAL DIR.',
        'PAROTIDECTOMIA SUPERFICIAL DIR.', 'PAROTIDECTOMIA SUPERFICAL ESQ.',
        'PAROTIDECTOMIA SUPERFICIAL DIR. COM PRESERVACAO DE NERVO FACIAL',
        'PAROTIDECTOMIA TOTAL + LINFADENECTOMIA RADICAL MODIFICADA UNILATERAL',
        'PAROTIDECTOMIA DIREITA TOTAL + RECONSTRUCAO COM RETALHO BILOBADO',
        'PAROTIDECTOMIA TOTAL + RECONTRUCAO COM RETALHO BILOBADO',
        'PAROTIDECTOMIA TOTAL + LINFADENECTOMIA RADICAL MODIFICADA',
        'PAROTIDECTOMIA TOTAL AMPLIADA A ESQ.',
        'PAROTIDECTOMIA TOTAL + RECONSTRUCAO COM RETALHO',
    ],
    'LARINGECTOMIA TOTAL': [
        'LARINGECTOMIA TOTAL', 'LARINGECTOMIA TOTAL + EC II-IV BILATERAL',
        'LARINGECTOMIA TOTAL + LINFADENECTOMIA RADICAL MODIFICADA BILATERAL',
        'LARINGECTOMIA TOTAL DE RESGATE',
        'LARINGECTOMIA TOTAL + EC UU A V BILATERAL',
    ],
    'LARINGECTOMIA PARCIAL': [
        'LARINGECTOMIA PARCIAL', 'LARINGECTOMIA PARCIAL + TQT',
        'LARINGECTOMIA PARCIAL FRONTO-LATERAL DIR.',
        'LARINGECTOMIA PARCIAL + TRAQUEOSTOMIA',
        'LARINGECTOMIA SUPRA CRICODEA', 'LARINGECTOMIA SUPRA CRICODEA + CHEP',
        'LARINGECTOMIA SUPRA CRIXOIDEA',
    ],
    'RESSECCAO DE TUMOR DE LABIO': [
        'RESSECCAO DE LESAO DE LABIO INFERIOR', 'RESSECCAO DE LESAO EM LABIO INFERIOR',
        'RESSECCAO EM CUNHA DE LABIO', 'RESSECÇÃO DE LESÃO DE LABIO INFERIOR',
        'RESSECCAO DE LESAO EM LABIO INFERIOR + RECONSTRUCAO COM RETALHO',
    ],
    'FARINGECTOMIA PARCIAL': [
        'FARINGECTOMIA PARCIAL', 'FARINGECTOMIA PARCIAL ESQ.', 'FARINGECTOMIA TOTAL',
        'FARINGECTOMIA + EC II A V A DIREITA', 'BUCOFARINGECTOMIA ESQ.',
        'BUCOFARINGECTOMIA AMPLIADA PARA BASE DE LINGUA',
        'FARINGOLARINGECTOMIA PARCIAL',
        'FARINGECTOMIA PARCIAL (AMIGDALECTOMIA ESQ.)',
    ],
    'SEQUESTRECTOMIA': ['MANDIBULECTOMIA PARCIAL', 'MANDIBULESTOMIA PARCIAL', 'SEQUESTRECTOMIA', 'RESSECCAO DE OSTEORADIONECROSE'],
    'RESSECÇÃO DE SUBMANDIBULAR': ['RESSECCAO DE GLANDULA SUBMANDIBULAR', 'SUBMANDIBULECTOMIA'],
    'RESSECÇÃO DE TUMOR DE PELE': [
        'EXCISAO E SUTURA', 'EXCISAO E SUTURA DE LESAO NA PELE',
        'EXCISÃO E SUTURA DE LESÃO NA PELE', 'RESSECCAO DE LESOES DE PELE',
        'BIOPSIA DE PELE E PARTES MOLES', 'BIOPSIA DE LESAO DE PARTES MOLES',
        'RESSECÇAO DE LESÃO DE PELE', 'RESSECCAO DE LESAO CROSTOSA NASAL',
        'EXTIRPACAO MULTIPLA DE LESAO DA PELE', 'AMPLIACAO DE MARGEM',
    ],
    'ESVAZIAMENTO CERVICAL': [
        'ESVAZIAMENTO CERVICAL', 'ESVAZIAMENTO CERVICAL ESQ.', 'ESVAZIAMENTO CERVICAL DIR.',
        'ESVAZIAMENTO CERVICAL UNILATERAL', 'ESVAZIAMENTO CERVICAL RECORRENCIAL BILATERAL',
        'ESVAZIAMENTO CERVICAL II-IV BILATERAL', 'ESVAZIAMENTO CERVICAL NIVEIS II-VI',
        'ESVAZIAMENTO CERVICAL RADICAL', 'ESVAZIAMENTO RADICAL MODIFICADO',
        'ESVAZIAMENTO CERVICAL RADICAL MODIFICADO', 'ESVAZIAMENTO CERVICAL SUPRA-OMOHIOIDEO',
        'LINFADENECTOMIA', 'LINFADENECTOMIA CERVICAL', 'LINFADENECTOMIA RADICAL',
    ],
    'BIOPSIA': ['BIOPSIA', 'BIOPSIA CERVICAL COM AGULHA DE TRUCUT'],
    'REOPERAÇÃO': ['REENXERTO', 'REOPERAÇÃO', 'RETORNO AO CENTRO CIRURGICO'],
    'RECONSTRUCAO COM RETALHO': ['RECONSTRUCAO COM RETALHO', 'RETALHO MIOCUTANEO', 'RECONSTRUÇÃO'],
    'BENIGNO': ['BENIGNO', 'EXCISE DE CISTO', 'DRENAGEM'],
}

# =====================================================
# FUNÇÕES DE LIMPEZA E NORMALIZAÇÃO
# =====================================================
def normalizar_medico(nome, lista_validos, dados=None, idx=None, perguntar=True):
    """Normaliza nome de médico com fuzzy matching."""
    global MEDICOS_VALIDOS
    
    if pd.isna(nome):
        return None
    
    nome_upper = str(nome).upper().strip()
    
    # Verifica se contém algum nome válido
    for valido in lista_validos:
        if valido.split()[0] in nome_upper:
            return valido
    
    # Fuzzy match - buscar similares
    results = process.extract(nome_upper, lista_validos, scorer=fuzz.ratio, limit=5)
    similares = [r[0] for r in results if r[1] > 40]
    
    if not perguntar:
        return nome_upper
    
    # Mostrar opções
    print(f"\n⚠️  Médico não identificado: {nome}")
    
    if similares:
        print("   Médicos similares encontrados:")
        for i, sim in enumerate(similares, 1):
            print(f"      [{i}] {sim}")
        print("      [L] Ver lista completa")
        print("      [N] Novo médico")
        print("      [E] Excluir linha")
        resp = input("   Escolha: ").lower().strip()
        
        if resp in ['1', '2', '3', '4', '5']:
            opcao = int(resp) - 1
            if opcao < len(similares):
                return similares[opcao]
        elif resp == 'e':
            return '__EXCLUIR__'
        elif resp == 'l':
            pass  # Vai para lista completa
        elif resp == 'n':
            nome_novo = input("      Nome: ").strip().upper()
            crm_novo = input("      CRM: ").strip()
            nova_lista = salvar_medico(nome_novo, crm_novo)
            MEDICOS_VALIDOS = nova_lista
            return f"{nome_novo} (CRM - {crm_novo})"
        elif resp == 'c':
            return nome_upper
    else:
        print("   Nenhum similar encontrado.")
    
    # Mostrar lista completa de médicos
    print("\n   Lista de médicos cadastrados:")
    for i, med in enumerate(lista_validos, 1):
        print(f"      [{i}] {med}")
    print("      [N] Novo médico")
    print("      [E] Excluir linha")
    print("      [C] Manter como está")
    resp2 = input("   Escolha: ").lower().strip()
    
    if resp2 == 'e':
        return '__EXCLUIR__'
    elif resp2 == 'n':
        nome_novo = input("      Nome: ").strip().upper()
        crm_novo = input("      CRM: ").strip()
        nova_lista = salvar_medico(nome_novo, crm_novo)
        MEDICOS_VALIDOS = nova_lista
        return f"{nome_novo} (CRM - {crm_novo})"
    elif resp2 == 'c':
        return nome_upper
    elif resp2.isdigit():
        idx = int(resp2) - 1
        if 0 <= idx < len(lista_validos):
            return lista_validos[idx]
    
    return nome_upper

# Carregar mapeamentos aprendidos
MAPEAMENTOS = {}
if os.path.exists(ARQ_MAPEAamentos):
    df_map = pd.read_excel(ARQ_MAPEAamentos)
    for _, row in df_map.iterrows():
        MAPEAMENTOS[str(row['CIRURGIA_ORIGINAL']).upper()] = str(row['GRUPO']).upper()

def mapear_grupo_fuzzy(cirurgia):
    """Mapeia cirurgia para grupo com fuzzy matching."""
    if pd.isna(cirurgia):
        return 'OUTRAS'
    
    cirurgia = str(cirurgia).upper().strip()
    
    # Verificar mapeamentos aprendidos
    if cirurgia in MAPEAMENTOS:
        return MAPEAMENTOS[cirurgia]
    
    # Fuzzy match com agrupamentos
    for grupo, lista in agrupamentos.items():
        result = process.extractOne(cirurgia, lista, scorer=fuzz.ratio)
        if result and result[1] > 80:
            return grupo
    return 'OUTRAS'

def salvar_mapeamento(cirurgia_original, grupo_novo):
    """Salva novo mapeamento aprendido."""
    MAPEAMENTOS[cirurgia_original.upper()] = grupo_novo.upper()
    
    novo = pd.DataFrame({'CIRURGIA_ORIGINAL': [cirurgia_original], 'GRUPO': [grupo_novo]})
    
    if os.path.exists(ARQ_MAPEAamentos):
        df_existente = pd.read_excel(ARQ_MAPEAamentos)
        df_novo = pd.concat([df_existente, novo], ignore_index=True)
    else:
        df_novo = novo
    
    df_novo.to_excel(ARQ_MAPEAamentos, index=False)

def limpar_dados(df):
    """Aplica todas as limpezas no DataFrame."""
    # Limpar nomes de colunas
    df.columns = df.columns.str.strip()
    
    # FILTRO: Selecionar apenas clínica CP
    colunas_possiveis = ['CLINICA', 'clinica', 'E']
    col_clinica = None
    for col in df.columns:
        if col.upper() in ['CLINICA', 'E'] or 'CLINICA' in col.upper():
            col_clinica = col
            break
    
    if col_clinica:
        before = len(df)
        df = df[df[col_clinica].astype(str).str.upper().str.strip() == 'CP']
        print(f"   🔍 Filtrado {before - len(df)} registros de outra clínica (CP: {len(df)})")
    
    # Normalizar texto
    for col in ['CHEFE', 'RESIDENTE', 'ANESTESISTA', 'CIRCULANTE']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()
    
    if 'CIRURGIA' in df.columns:
        df['CIRURGIA'] = df['CIRURGIA'].astype(str).str.strip().str.upper()
    
    # Normalizar médicos (com interação)
    if 'CHEFE' in df.columns:
        Chefe_normalizado = []
        for idx, nome in df['CHEFE'].items():
            resultado = normalizar_medico(nome, MEDICOS_VALIDOS, df, idx, perguntar=True)
            if resultado == '__EXCLUIR__':
                Chefe_normalizado.append('__EXCLUIR__')
            else:
                Chefe_normalizado.append(resultado)
        
        # Remover linhas excluídas
        df['CHEFE'] = Chefe_normalizado
        df = df[df['CHEFE'] != '__EXCLUIR__']
    
    # Mapear grupos de cirurgia
    if 'CIRURGIA' in df.columns:
        df['CIRURGIA_GRUPO'] = df['CIRURGIA'].apply(mapear_grupo_fuzzy)
    
    # Filtrar médicos fora da CCP (opcional)
    MEDICOS_EXCLUIR = ['NIKKEI TAMURA', 'CARLOS ELIAS', 'RODRIGO MACEDO']
    if 'CHEFE' in df.columns:
        df = df[~df['CHEFE'].str.contains('|'.join(MEDICOS_EXCLUIR), na=False)]
    
    return df

# =====================================================
# FUNÇÃO PRINCIPAL
# =====================================================
def processar_mes(mes_num, ano, arquivo_novo=None, arquivo_saida=None):
    """
    Processa dados de um mês específico.
    
    Args:
        mes_num: Número do mês (1-12)
        ano: Ano (ex: 2025)
        arquivo_novo: Caminho para arquivo Excel do mês
        arquivo_saida: Arquivo de saída consolidado
    """
    # Arquivo de saída na pasta de trabalho
    if arquivo_saida is None:
        arquivo_saida = os.path.join(ARQUIVOS_PASTA, 'cirurgias_cp_MM.xlsx')
    
    from pathlib import Path
    
    # Pedir arquivo se não informado
    if arquivo_novo is None:
        arquivo_novo = input(f"Digite o caminho do arquivo Excel ({mes_nome} {ano}): ").strip().strip('"')
    
    if not os.path.exists(arquivo_novo):
        print(f"❌ Arquivo não encontrado: {arquivo_novo}")
        return
    
    print(f"\n📂 Lendo {arquivo_novo}...")
    df_novo = pd.read_excel(arquivo_novo)
    print(f"   {len(df_novo)} registros lidos.")
    
    # Limpar dados
    print("🧹 Limpando dados...")
    df_novo = limpar_dados(df_novo)
    
    # Converter datas
    if 'DATA' in df_novo.columns:
        df_novo['DATA'] = pd.to_datetime(df_novo['DATA'], errors='coerce')
    
    # Calcular duração
    if 'INICIO' in df_novo.columns and 'FIM' in df_novo.columns:
        df_novo['INICIO'] = pd.to_datetime(df_novo['DATA'].astype(str) + ' ' + df_novo['INICIO'].astype(str), errors='coerce')
        df_novo['FIM'] = pd.to_datetime(df_novo['DATA'].astype(str) + ' ' + df_novo['FIM'].astype(str), errors='coerce')
        df_novo['DURACAO_MIN'] = (df_novo['FIM'] - df_novo['INICIO']).dt.total_seconds() / 60
        df_novo['DURACAO_HORAS'] = df_novo['DURACAO_MIN'] / 60
    
    # Verificar se arquivo consolidado existe
    mes_nome = MESES.get(mes_num, '')
    
    if os.path.exists(arquivo_saida):
        df_total = pd.read_excel(arquivo_saida)
        
        # Converter datas para comparação
        df_total['DATA'] = pd.to_datetime(df_total['DATA'], errors='coerce')
        
        # Verificar se mês já existe
        mes_str = f"{ano}-{str(mes_num).zfill(2)}"
        meses_existentes = df_total['DATA'].dt.to_period('M').astype(str).unique()
        
        if mes_str in meses_existentes:
            # Verificar pacientes duplicados (por MV)
            pacientes_mes_novo = set(df_novo['MV'].unique())
            pacientes_mes_existente = set(df_total[df_total['DATA'].dt.to_period('M').astype(str) == mes_str]['MV'].unique())
            duplicados = pacientes_mes_novo & pacientes_mes_existente
            
            if duplicados:
                qtd_novo = len(df_novo)
                qtd_geral = len(pacientes_mes_existente)
                qtd_duplicados = len(duplicados)
                print(f"\n⚠️  {mes_nome} {ano} já existe!")
                print(f"   Novos: {qtd_novo} | Existentes: {qtd_geral} | Duplicados (por MV): {qtd_duplicados}")
                print(f"\nEscolha:")
                print(f"   [R] Substituir todos os dados deste mês")
                print(f"   [A] Abortar過程")
                print(f"   [C] Continuar (manter existentes + adicionar novos)")
                resposta = input("   Opção: ").lower().strip()
                
                if resposta == 'a':
                    print("❌ Cancelado.")
                    return None
                elif resposta == 'r':
                    # Remover mês antigo
                    df_total = df_total[~df_total['DATA'].dt.to_period('M').astype(str).eq(mes_str)]
                    print(f"   Mês antigo removido.")
                elif resposta == 'c':
                    # Manter existentes, remover pacientes duplicados do novo
                    df_novo = df_novo[~df_novo['MV'].isin(duplicados)]
                    print(f"   {qtd_duplicados} pacientes duplicados removidos do novo arquivo.")
                else:
                    print("❌ Opção inválida. Abortando.")
                    return None
        
        print(f"📂 Atualizando {arquivo_saida}...")
        df_final = pd.concat([df_total, df_novo], ignore_index=True)
    else:
        print("📂 Criando novo arquivo consolidado...")
        df_final = df_novo
    
    # Selecionar colunas finais
    COLUNAS_FINAIS = [
        'DATA', 'MV', 'CHEFE', 'CIRURGIA', 'ANEST', 'ANESTESISTA',
        'INICIO', 'FIM', 'CIRCULANTE', 'CIRURGIA_GRUPO',
        'DURACAO_MIN', 'DURACAO_HORAS', 'GRUPO_MESTRE',
        'COMPLICACAO', 'QUAL (1)', 'QUAL (2)',
        'HORA_ALTA', 'TEMPO_INTERNACAO_HORAS', 'TEMPO_INTERNACAO_DIAS',
        'REOPERACAO', 'OBITO', 'DATA_ÓBITO', 'REINTERNACAO_NAO_PROGRAMADA'
    ]
    
    colunas_presentes = [c for c in COLUNAS_FINAIS if c in df_final.columns]
    df_final = df_final[colunas_presentes]
    
    # Salvar
    df_final.to_excel(arquivo_saida, index=False)
    print(f"✅ Salvo: {arquivo_saida}")
    print(f"   Total de registros: {len(df_final)}")
    
    # Estatísticas
    print("\n📊 Resumo:")
    if 'CIRURGIA_GRUPO' in df_final.columns:
        print(df_final['CIRURGIA_GRUPO'].value_counts())
    
    # Casos não mapeados
    outras = df_novo[df_novo['CIRURGIA_GRUPO'] == 'OUTRAS']['CIRURGIA'].unique()
    if len(outras) > 0:
        print(f"\n⚠️  {len(outras)} cirurgias não mapeadas (OUTRAS):")
        for c in outras[:5]:
            print(f"   - {c}")
        if len(outras) > 5:
            print(f"   ... e mais {len(outras) - 5}")
        
        # Perguntar se quer mapear
        print("\n" + "="*40)
        print("[D] Mapear cada cirurgia (paciente por paciente)")
        print("[S] Salvar e sair agora (OUTRAS)")
        print("[K] Manter como está")
        print("[C] Cancelar tudo")
        print("="*40)
        resp = input("Escolha: ").strip().lower()
        
        print(f"\n>>> Você escolheu: '{resp}'")
        
        if resp == 'd':
            print("\n>> ENTROU no modo mapeamento!")
            print(f">> Total de não mapeadas: {len(outras)}")
            print("\n📝 Mapeando cirurgias uma por uma...")
            for cir in outras:
                #Buscar dados do paciente
                try:
                    paciente = df_novo[df_novo['CIRURGIA'] == cir].iloc[0]
                    mv = paciente.get('MV', '?')
                    data = paciente.get('DATA', '?')
                    chefe = paciente.get('CHEFE', '?')
                except:
                    mv, data, chefe = '?', '?', '?'
                
                # Buscar melhores matches via fuzzy
                todos_grupos = list(agrupamentos.keys())
                results = process.extract(cir, todos_grupos, scorer=fuzz.ratio, limit=5)
                matches = [(r[0], r[1]) for r in results if r[1] > 30]
                
                print(f"\n   ========== PACIENTE {mv} | {data} ==========")
                print(f"   Cirurgia: {cir}")
                print(f"   Chefe: {chefe}")
                
                if matches:
                    print("\n   Melhores matches sugeridos:")
                    for i, (grp, score) in enumerate(matches, 1):
                        print(f"      [{i}] {grp} ({score}%)")
                
                print("\n   Grupos disponíveis:")
                for i, grupo in enumerate(list(agrupamentos.keys())[:10], 1):
                    print(f"      [{i+10}] {grupo}")
                if len(agrupamentos) > 10:
                    print("      [...mais]")
                print("   [B] Buscar melhor opção (fuzzy)")
                print("   [N] Novo grupo")
                print("   [K] Manter como está (OUTRAS)")
                print("   [C] Cancelar tudo")
                opc = input("   Escolha: ").lower().strip()
                
                if opc == 'b':
                    if matches:
                        melhor = matches[0][0]
                        print(f"   ➜ Melhor opção: {melhor} ({matches[0][1]}%)")
                        conf = input(f"   Confirmar? (s/{melhor}/n): ").lower().strip()
                        if conf == 's':
                            grupo_novo = melhor
                        elif conf.isdigit():
                            idx = int(conf) - 1
                            if 0 <= idx < len(agrupamentos):
                                grupo_novo = list(agrupamentos.keys())[idx]
                            else:
                                continue
                        else:
                            grupo_novo = conf.upper()
                    else:
                        grupo_novo = input("   Grupo: ").strip().upper()
                    salvar_mapeamento(cir, grupo_novo)
                    df_novo.loc[df_novo['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                    df_final.loc[df_final['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                    print(f"      ➜ Mapeado para: {grupo_novo}")
                elif opc.isdigit():
                    idx = int(opc) - 1
                    if 0 <= idx < len(agrupamentos):
                        grupo_novo = list(agrupamentos.keys())[idx]
                        salvar_mapeamento(cir, grupo_novo)
                        df_novo.loc[df_novo['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                        df_final.loc[df_final['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                        print(f"      ➜ Mapeado para: {grupo_novo}")
                elif opc == 'n':
                    grupo_novo = input("      Nome do novo grupo: ").strip().upper()
                    salvar_mapeamento(cir, grupo_novo)
                    df_novo.loc[df_novo['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                    df_final.loc[df_final['CIRURGIA'] == cir, 'CIRURGIA_GRUPO'] = grupo_novo
                    print(f"      ➜ Novo grupo: {grupo_novo}")
                elif opc in ['k', 'c']:
                    pass  # Mantém como OUTRAS
                else:
                    pass
            
            # Salvar novamente com mapeamentos
            if 'CIRURGIA_GRUPO' in df_novo.columns:
                df_final = pd.concat([df_total, df_novo], ignore_index=True)
                df_final.to_excel(arquivo_saida, index=False)
                print(f"\n✅ Salvo com mapeamentos atualizados!")
        
        elif resp in ['s', 'k']:
            pass  # Salvar e sair
        
        elif resp == 'c':
            print("❌ Cancelado.")
            return None
    
    return df_final


# =====================================================
# EXECUÇÃO INTERATIVA
MESES = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

ARQUIVOS_PASTA = r"F:\documentos francisco\Trabalho\DataScience2\Morbimortalidade CCP IAVC\cp - organizador planilhas para MM"

# =====================================================
if __name__ == "__main__":
    arquivo_saida = os.path.join(ARQUIVOS_PASTA, 'cirurgias_cp_MM.xlsx')
    
    print("=" * 50)
    print("📋 PROCESSADOR DE CIRURGIAS CCP")
    print("=" * 50)
    
# Verificar dados existentes
    arquivo_saida = os.path.join(ARQUIVOS_PASTA, 'cirurgias_cp_MM.xlsx')

    
    if os.path.exists(arquivo_saida):
        df_existente = pd.read_excel(arquivo_saida)
        
        # Detectar coluna de data (primeira coluna ou nome 'DATA')
        col_data = None
        for col in df_existente.columns:
            if 'DATA' in col.upper() or col == 'A':
                col_data = col
                break
        
        if col_data:
            # Converter data (tentarDD/MM/YYYY primeiro)
            try:
                df_existente[col_data] = pd.to_datetime(df_existente[col_data], dayfirst=True, errors='coerce')
            except:
                df_existente[col_data] = pd.to_datetime(df_existente[col_data], errors='coerce')
            
            # Verificar datas válidas
            datas_validas = df_existente[col_data].dropna()
            if len(datas_validas) > 0:
                primeiro = datas_validas.min()
                ultimo = datas_validas.max()
                meses_existentes = sorted(datas_validas.dt.to_period('M').unique())
                
                print(f"\n📁 Dados consolidados: {len(datas_validas)} cirurgias")
                print(f"   Período: {primeiro.strftime('%d/%m/%Y')} até {ultimo.strftime('%d/%m/%Y')}")
                print(f"   Meses covered: {len(meses_existentes)}")
    else:
        print("\n📁 Nenhum dado consolidado encontrado.")
    
    print("\nMeses: 1=Jan, 2=Fev, 3=Mar, 4=Abr, 5=Mai, 6=Jun")
    print("       7=Jul, 8=Ago, 9=Set, 10=Out, 11=Nov, 12=Dez")
    
    try:
        mes_num = int(input("\nMês (número 1-12): "))
        ano = int(input("Ano (ex: 2025): "))
    except ValueError:
        print("❌ Mês ou ano inválido!")
        exit()
    
    if mes_num not in MESES:
        print("❌ Mês inválido! Use 1-12.")
        exit()
    
    mes_nome = MESES[mes_num]
    arquivo_encontrado = None
    
    # Buscar arquivo na pasta
    for f in os.listdir(ARQUIVOS_PASTA):
        if f.startswith(mes_nome) and f.endswith('.xlsx'):
            arquivo_encontrado = os.path.join(ARQUIVOS_PASTA, f)
            break
    
    if arquivo_encontrado:
        print(f"\n📂 Arquivo encontrado: {os.path.basename(arquivo_encontrado)}")
        confirmar = input("Confirmar? (s/n): ").lower().strip()
        if confirmar != 's':
            print("❌ Cancelado.")
            exit()
    else:
        print(f"\n❌ Arquivo '{mes_nome} {ano}.xlsx' não encontrado na pasta!")
        print(f"\n   [M] Informar caminho manualmente")
        print("   [L] Listar arquivos disponíveis")
        print("   [C] Cancelar")
        resp = input("   Escolha: ").lower().strip()
        
        if resp == 'l':
            print("\n   Arquivos na pasta:")
            arquivos_xlsx = [f for f in sorted(os.listdir(ARQUIVOS_PASTA)) 
                         if f.endswith('.xlsx') and not f.startswith('~') 
                         and not f.startswith('cirurgias_cp')]
            for i, f in enumerate(arquivos_xlsx, 1):
                print(f"   [{i}] {f}")
            print("\n   [N] Escolher por número")
            arquivo_escollha = input("   Arquivo: ").strip()
            if arquivo_escollha.isdigit():
                idx = int(arquivo_escollha) - 1
                if 0 <= idx < len(arquivos_xlsx):
                    arquivo_encontrado = os.path.join(ARQUIVOS_PASTA, arquivos_xlsx[idx])
                    print(f"   ✅ Arquivo seleccionado: {arquivos_xlsx[idx]}")
                else:
                    print("   ❌ Número inválido!")
                    exit()
            else:
                print("   ❌ Número inválido!")
                exit()
        elif resp == 'm':
            arquivo_manual = input("   Caminho do arquivo: ").strip().strip('"')
            if os.path.exists(arquivo_manual):
                arquivo_encontrado = arquivo_manual
                print(f"   ✅ Arquivo confirmado!")
            else:
                print(f"   ❌ Arquivo não existe!")
                exit()
        else:
            print("❌ Cancelado.")
            exit()
    
    processar_mes(mes_num, ano, arquivo_encontrado)