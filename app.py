import streamlit as st
import os
import tempfile
import base64
from curl_cffi import requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Relação de Autoridades Municipais", page_icon="🗳️", layout="centered")

# --- CUSTOMIZAÇÃO DE DESIGN (TELA ESCURA MODERNA) ---
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] {
        background-color: #0E1117; 
    }
    [data-testid="stHeader"] {
        background-color: rgba(0,0,0,0);
    }
    h1 {
        color: #4DB8FF !important;
        font-weight: 700;
    }
    .stMarkdown p, label {
        color: #E0E6ED !important;
        font-size: 16px;
    }
    .stButton>button {
        background-color: #1F618D !important;
        border: 1px solid #2980B9 !important;
        border-radius: 8px;
    }
    .stButton>button * {
        color: #FFFFFF !important;
    }
    .stButton>button:hover {
        background-color: #2980B9 !important;
        border-color: #4DB8FF !important;
    }
    [data-testid="stAlert"] {
        background-color: #0E3B21 !important;
        border: 1px solid #145A32;
    }
    [data-testid="stAlert"] * {
        color: #D1FAE5 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- TÍTULOS E DESCRIÇÃO ---
st.title("🗳️ Relação de autoridades municipais")
st.markdown("*(Informações extraídas da base de dados do TSE referência eleições de 2024)*")
st.markdown("Selecione o município desejado na lista abaixo para extrair a relação oficial de representantes e gerar automaticamente o documento formatado.")

# Parâmetros Fixos
ANO = "2024"
ID_ELEICAO = "2045202024"
ESTADO = "GO"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

@st.cache_data
def buscar_municipios():
    url_municipios = f"https://divulgacandcontas.tse.jus.br/divulga/rest/v1/eleicao/buscar/{ESTADO}/{ID_ELEICAO}/municipios"
    try:
        resp = requests.get(url_municipios, impersonate="chrome120", timeout=20)
        if resp.status_code == 200:
            return {m.get('nome').upper(): m.get('codigo') for m in resp.json().get("municipios", [])}
    except:
        pass
    return {}

municipios_go = buscar_municipios()

if municipios_go:
    cidade_escolhida = st.selectbox("Selecione o Município:", [""] + sorted(list(municipios_go.keys())))

    if cidade_escolhida:
        if st.button(f"🔍 Gerar Relatório de {cidade_escolhida}"):
            cod_municipio = municipios_go[cidade_escolhida]
            
            with st.spinner(f"Buscando e processando dados para {cidade_escolhida}... Isso pode levar alguns segundos."):
                cargos_para_buscar = [("11", "Prefeito"), ("13", "Vereador")]
                
                temp_dir = tempfile.TemporaryDirectory()
                pasta_fotos = temp_dir.name
                eleitos = []
                
                for cod_cargo, nome_cargo in cargos_para_buscar:
                    url_lista = f"https://divulgacandcontas.tse.jus.br/divulga/rest/v1/candidatura/listar/{ANO}/{cod_municipio}/{ID_ELEICAO}/{cod_cargo}/candidatos"
                    try:
                        resp = requests.get(url_lista, impersonate="chrome120", timeout=30)
                        if resp.status_code == 200:
                            for cand in resp.json().get("candidatos", []):
                                if "Eleito" in cand.get("descricaoTotalizacao", ""):
                                    cand_id = cand["id"]
                                    url_detalhe = f"https://divulgacandcontas.tse.jus.br/divulga/rest/v1/candidatura/buscar/{ANO}/{cod_municipio}/{ID_ELEICAO}/candidato/{cand_id}"
                                    
                                    try:
                                        detalhe_resp = requests.get(url_detalhe, impersonate="chrome120", timeout=20)
                                        detalhe = detalhe_resp.json()
                                        
                                        nome_urna = detalhe.get("nomeUrna", cand.get("nomeUrna", "Sem Nome"))
                                        nome_completo = detalhe.get("nomeCompleto", cand.get("nomeCompleto", ""))
                                        partido = detalhe.get("partido", {}).get("sigla", cand.get("partido", {}).get("sigla", ""))
                                        foto_url = detalhe.get("fotoUrl")
                                        caminho_foto = os.path.join(pasta_fotos, f"{cand_id}.jpg")
                                        
                                        if foto_url:
                                            img_resp = requests.get(foto_url, impersonate="chrome120", timeout=20)
                                            if img_resp.status_code == 200:
                                                with open(caminho_foto, "wb") as f:
                                                    f.write(img_resp.content)
                                                    
                                        eleitos.append({
                                            "nome_urna": nome_urna, "nome_completo": nome_completo,
                                            "cargo": nome_cargo, "partido": partido,
                                            "foto_local": caminho_foto if os.path.exists(caminho_foto) else None
                                        })
                                        
                                        if cod_cargo == "11":
                                            for v in detalhe.get("vices", []):
                                                nome_v_urna = v.get("nomeUrna", v.get("nm_URNA", "Sem Nome"))
                                                nome_v_completo = v.get("nomeCompleto", v.get("nm_CANDIDATO", ""))
                                                partido_v = v.get("partido", "")
                                                if isinstance(partido_v, dict): partido_v = partido_v.get("sigla", "")
                                                foto_v_url = v.get("urlFoto") or v.get("fotoUrl")
                                                caminho_foto_v = os.path.join(pasta_fotos, f"vice_{cand_id}.jpg")
                                                
                                                if foto_v_url:
                                                    try:
                                                        img_resp_v = requests.get(foto_v_url, impersonate="chrome120", timeout=20)
                                                        if img_resp_v.status_code == 200:
                                                            with open(caminho_foto_v, "wb") as f:
                                                                f.write(img_resp_v.content)
                                                    except: pass
                                                eleitos.append({
                                                    "nome_urna": nome_v_urna, "nome_completo": nome_v_completo,
                                                    "cargo": "Vice-Prefeito", "partido": partido_v,
                                                    "foto_local": caminho_foto_v if os.path.exists(caminho_foto_v) else None
                                                })
                                    except: pass
                    except: pass

                if not eleitos:
                    st.error("Não foi possível encontrar eleitos para esta cidade. Tente novamente.")
                else:
                    # GERAÇÃO DO WORD PROFISSIONAL
                    doc = Document()
                    
                    # Ajuste das Margens
                    for section in doc.sections:
                        section.top_margin = Inches(0.6)
                        section.bottom_margin = Inches(0.6)
                        section.left_margin = Inches(0.8)
                        section.right_margin = Inches(0.8)

                    # Cabeçalho do Documento
                    title_p = doc.add_paragraph()
                    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    title_run = title_p.add_run(f'RELAÇÃO DE AUTORIDADES MUNICIPAIS\n{cidade_escolhida.upper()} - GO')
                    title_run.bold = True
                    title_run.font.size = Pt(16)
                    title_run.font.color.rgb = RGBColor(31, 73, 125) # Azul Escuro
                    
                    subtitle_p = doc.add_paragraph()
                    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    subtitle_run = subtitle_p.add_run('Gestão 2025-2028 | Dados extraídos do TSE')
                    subtitle_run.font.size = Pt(10)
                    subtitle_run.font.color.rgb = RGBColor(128, 128, 128) # Cinza
                    
                    doc.add_paragraph() # Espaço extra

                    # Configuração da Tabela
                    table = doc.add_table(rows=1, cols=3)
                    table.style = 'Table Grid'
                    table.autofit = False

                    # Largura das Colunas
                    widths = (Inches(1.2), Inches(3.8), Inches(1.5))
                    for row in table.rows:
                        for idx, width in enumerate(widths):
                            row.cells[idx].width = width

                    # Formatação da Linha de Cabeçalho
                    hdr_cells = table.rows[0].cells
                    cabecalhos = ['FOTO', 'DADOS DO ELEITO', 'PARTIDO']
                    for i in range(3):
                        hdr_cells[i].text = cabecalhos[i]
                        set_cell_background(hdr_cells[i], '1F497D') # Fundo Azul Escuro
                        hdr_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                        for paragraph in hdr_cells[i].paragraphs:
                            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in paragraph.runs: 
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255) # Texto Branco

                    # Preenchimento das Linhas (com efeito Zebra)
                    for index, v in enumerate(eleitos):
                        row_cells = table.add_row().cells
                        
                        # Mantém as larguras
                        for idx, width in enumerate(widths):
                            row_cells[idx].width = width

                        # Cor de fundo (Branco par, Cinza ultra claro ímpar)
                        bg_color = 'FFFFFF' if index % 2 == 0 else 'F8F9FA'
                        for cell in row_cells:
                            set_cell_background(cell, bg_color)
                            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER # Centraliza no meio

                        # Célula Foto
                        cell_foto = row_cells[0]
                        p_foto = cell_foto.paragraphs[0]
                        p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if v.get("foto_local") and os.path.exists(v.get("foto_local")):
                            p_foto.add_run().add_picture(v["foto_local"], width=Inches(1.0))
                        else:
                            run_nofoto = p_foto.add_run("Sem foto")
                            run_nofoto.font.color.rgb = RGBColor(160, 160, 160)

                        # Célula Nome e Cargo
                        cell_nome = row_cells[1]
                        p_nome = cell_nome.paragraphs[0]
                        p_nome.paragraph_format.space_after = Pt(2)
                        
                        run_urna = p_nome.add_run(f'{v["nome_urna"].upper()}\n')
                        run_urna.font.bold = True
                        run_urna.font.size = Pt(11)
                        
                        run_comp = p_nome.add_run(f'{v["nome_completo"].title()}\n')
                        run_comp.font.size = Pt(9)
                        run_comp.font.color.rgb = RGBColor(89, 89, 89)
                        
                        run_cargo = p_nome.add_run(v['cargo'].upper())
                        run_cargo.font.bold = True
                        run_cargo.font.size = Pt(9)
                        run_cargo.font.color.rgb = RGBColor(31, 73, 125)

                        # Célula Partido
                        cell_partido = row_cells[2]
                        p_partido = cell_partido.paragraphs[0]
                        p_partido.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run_part = p_partido.add_run(v['partido'])
                        run_part.font.bold = True
                        run_part.font.size = Pt(11)

                    nome_arquivo = f'Autoridades_{cidade_escolhida.replace(" ", "_")}_GO.docx'
                    caminho_final = os.path.join(pasta_fotos, nome_arquivo)
                    doc.save(caminho_final)
                    
                    st.success("✅ Relatório gerado com sucesso!")
                    
                    with open(caminho_final, "rb") as file:
                        st.download_button(
                            label=f"⬇️ Baixar Relatório de {cidade_escolhida.title()}",
                            data=file,
                            file_name=nome_arquivo,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
else:
    st.error("Falha ao se conectar com os servidores para buscar as cidades.")
