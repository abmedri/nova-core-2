# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. - NOVA CORE 2 (MÓDULO DE CONFERÊNCIA)
# ====================================================================================

import os
import io
import pandas as pd
import streamlit as st

# Dependências para planilhas e PDFs
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

def processar_arquivo_upload(arquivo_carregado):
    nome = arquivo_carregado.name.lower()
    try:
        if nome.endswith('.csv'):
            for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                for sep in [';', ',']:
                    try:
                        arquivo_carregado.seek(0)
                        df = pd.read_csv(arquivo_carregado, encoding=enc, sep=sep, low_memory=False)
                        if df.shape[1] > 1:
                            return df, "Inventário processado com sucesso."
                    except Exception:
                        continue
            return None, "Erro: Não foi possível decodificar o ficheiro CSV."
        elif nome.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(arquivo_carregado)
            return df, "Excel processado com sucesso."
        else:
            return None, "Formato não suportado."
    except Exception as e:
        return None, "Erro ao processar ficheiro: " + str(e)

def limpar_nome_local(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ====================================================================================
# CANVAS NUMERADO (RODAPÉ OFICIAL)
# ====================================================================================
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()
    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor('#64748b'))
        
        largura_pagina, altura_pagina = self._pagesize
        margem = 10
        
        self.setStrokeColor(colors.HexColor('#cbd5e1'))
        self.setLineWidth(0.5)
        self.line(margem, 22, largura_pagina - margem, 22)
        
        caminho_brasao = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
        if caminho_brasao:
            try:
                self.drawImage(caminho_brasao, margem, 6, width=45, height=13, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
                
        texto_esquerda = "GRUPO KMB // Auditoria Patrimonial • Exatidão • Controle • Eficiência"
        self.drawString(margem + 50, 10, texto_esquerda)
        
        texto_pagina = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(largura_pagina - margem, 10, texto_pagina)
        self.restoreState()

# ====================================================================================
# GERADOR DE PDF DE CONFERÊNCIA OTIMIZADO
# ====================================================================================
def gerar_pdf_conferencia_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    soma_comptar_por_produto = {}
    descricoes_comptar = {}
    
    for _, row in df_localizacao.iterrows():
        loc = limpar_nome_local(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, '')).strip().replace('.0', '')
        desc = str(row.get(col_desc_loc, '')).strip()
        try:
            qtd = int(float(row.get(col_qtd_loc, 0)))
        except Exception:
            qtd = 0
            
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
            soma_comptar_por_produto[cod] = 0
            descricoes_comptar[cod] = desc
        locais_por_produto[cod].append((loc, qtd))
        soma_comptar_por_produto[cod] += qtd
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    
    codigos_estoque_set = set(df_estoque[col_cod_est].astype(str).str.strip().str.replace('.0', '', regex=False).values)
    
    itens_filtrados_temp = []
    if tipo_filtro == "SEM_SALDO":
        for cod, locs in locais_por_produto.items():
            if cod not in codigos_estoque_set:
                desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
                soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
                itens_filtrados_temp.append((cod, desc_item, 0, True, soma_comptar_val, soma_comptar_val, locs))
    else:
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip().replace('.0', '')
            desc = str(row.get(col_desc_est, ''))
            try:
                est = int(float(row.get(col_est, 0)))
            except Exception:
                est = 0
            
            locs_prod = locais_por_produto.get(cod, [])
            tem_na_comptar = len(locs_prod) > 0
            soma_comptar_val = soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
            divergencia_val = soma_comptar_val - est
            
            incluir = False
            if tipo_filtro == "TODOS":
                incluir = True
            elif tipo_filtro == "ZERADOS":
                if not tem_na_comptar:
                    incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))
                
    itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1
    buffer = io.BytesIO()
    
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=10,
        leftMargin=10,
        topMargin=18,
        bottomMargin=25
    )
    
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#1F4E78'))
    
    estilo_th = ParagraphStyle('TH_Auto', fontName='Helvetica-Bold', fontSize=5.5, leading=6.5, textColor=colors.white, alignment=1)
    estilo_td_codigo = ParagraphStyle('TD_Codigo', fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_centro = ParagraphStyle('TD_Centro', fontName='Helvetica', fontSize=6.0, leading=7.5, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_esq = ParagraphStyle('TD_Esq', fontName='Helvetica', fontSize=7.0, leading=9, textColor=colors.HexColor('#111827'), alignment=0)
    
    story = []
    story.append(Paragraph(f"<b>GRUPO KMB // RELATÓRIO DE CONFERÊNCIA ({tipo_filtro})</b> — LOJA: {numero_loja_input} | DATA: {pd.Timestamp.now().strftime('%d/%m/%Y')}", estilo_titulo))
    story.append(Spacer(1, 3))
    
    headers_pdf = ["Código", "Descrição", "EST.", "COMP.", "DIVERG."]
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append("C.1" if i == 1 else str(int(i)))
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append(f"Loc.{int(i)}" if i == 1 else str(int(i)))
        
    dados_tabela = [[Paragraph(f"<nobr>{str(h).strip().replace('.0', '').replace(' ', '&nbsp;')}</nobr>", estilo_th) for h in headers_pdf]]
    
    for cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod in itens_filtrados_temp:
        cel_comptar = soma_comptar_val if tem_na_comptar else 0
        cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)
        
        cod_formatado = str(cod).strip().replace(" ", "&nbsp;")
        
        linha_celulas = [
            Paragraph(f"<nobr>{cod_formatado}</nobr>", estilo_td_codigo),
            Paragraph(str(desc), estilo_td_esq),
            Paragraph(f"<nobr>{est}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_comptar}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_divergencia}</nobr>", estilo_td_centro)
        ]
        
        if tem_na_comptar:
            contagens_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    val_tratado = str(int(float(locs_prod[i][1]))).replace('.', '&#46;')
                    contagens_vals.append(Paragraph(f"<nobr>{val_tratado}</nobr>", estilo_td_centro))
                else:
                    contagens_vals.append(Paragraph("", estilo_td_centro))
            
# ====================================================================================
# GERADOR DE PDF DE CONFERÊNCIA OTIMIZADO (LARGURAS AJUSTADAS PARA NÃO CORTAR CÓDIGOS)
# ====================================================================================
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()
    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor('#64748b'))
        
        largura_pagina, altura_pagina = self._pagesize
        margem = 10
        
        self.setStrokeColor(colors.HexColor('#cbd5e1'))
        self.setLineWidth(0.5)
        self.line(margem, 22, largura_pagina - margem, 22)
        
        caminho_brasao = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
        if caminho_brasao:
            try:
                self.drawImage(caminho_brasao, margem, 6, width=45, height=13, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
                
        texto_esquerda = "GRUPO KMB // Auditoria Patrimonial • Exatidão • Controle • Eficiência"
        self.drawString(margem + 50, 10, texto_esquerda)
        
        texto_pagina = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(largura_pagina - margem, 10, texto_pagina)
        self.restoreState()

def gerar_pdf_conferencia_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    soma_comptar_por_produto = {}
    descricoes_comptar = {}
    
    for _, row in df_localizacao.iterrows():
        loc = limpar_nome_local(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, '')).strip()
        desc = str(row.get(col_desc_loc, '')).strip()
        try:
            qtd = int(float(row.get(col_qtd_loc, 0)))
        except Exception:
            qtd = 0
            
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
            soma_comptar_por_produto[cod] = 0
            descricoes_comptar[cod] = desc
        locais_por_produto[cod].append((loc, qtd))
        soma_comptar_por_produto[cod] += qtd
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    codigos_estoque_set = set(df_estoque[col_cod_est].astype(str).str.strip().values)
    
    itens_filtrados_temp = []
    if tipo_filtro == "SEM_SALDO":
        for cod, locs in locais_por_produto.items():
            if cod not in codigos_estoque_set:
                desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
                soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
                itens_filtrados_temp.append((cod, desc_item, 0, True, soma_comptar_val, soma_comptar_val, locs))
    else:
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip()
            desc = str(row.get(col_desc_est, ''))
            try:
                est = int(float(row.get(col_est, 0)))
            except Exception:
                est = 0
            
            locs_prod = locais_por_produto.get(cod, [])
            tem_na_comptar = len(locs_prod) > 0
            soma_comptar_val = soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
            divergencia_val = soma_comptar_val - est
            
            incluir = False
            if tipo_filtro == "TODOS":
                incluir = True
            elif tipo_filtro == "ZERADOS":
                if not tem_na_comptar:
                    incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))
    itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1
    buffer = io.BytesIO()
    
    # 🛡️ Margens estreitas (10pt) para aproveitar ao máximo a largura da página paisagem
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=10,
        leftMargin=10,
        topMargin=18,
        bottomMargin=25
    )
    
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#1F4E78'))
    
    estilo_th = ParagraphStyle('TH_Auto', fontName='Helvetica-Bold', fontSize=5.5, leading=6.5, textColor=colors.white, alignment=1)
    estilo_td_codigo = ParagraphStyle('TD_Codigo', fontName='Helvetica-Bold', fontSize=6.5, leading=8, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_centro = ParagraphStyle('TD_Centro', fontName='Helvetica', fontSize=6.0, leading=7.5, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_esq = ParagraphStyle('TD_Esq', fontName='Helvetica', fontSize=7.0, leading=9, textColor=colors.HexColor('#111827'), alignment=0)
    
    story = []
    story.append(Paragraph(f"<b>GRUPO KMB // RELATÓRIO DE CONFERÊNCIA ({tipo_filtro})</b> — LOJA: {numero_loja_input} | DATA: {pd.Timestamp.now().strftime('%d/%m/%Y')}", estilo_titulo))
    story.append(Spacer(1, 3))
    
    headers_pdf = ["Código", "Descrição", "EST.", "COMP.", "DIVERG."]
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append("C.1" if i == 1 else str(int(i)))
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append(f"Loc.{int(i)}" if i == 1 else str(int(i)))
        
    dados_tabela = [[Paragraph(f"<nobr>{str(h).strip().replace('.0', '').replace(' ', '&nbsp;')}</nobr>", estilo_th) for h in headers_pdf]]
    
    for cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod in itens_filtrados_temp:
        cel_comptar = soma_comptar_val if tem_na_comptar else 0
        cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)
        
        cod_formatado = str(cod).strip().replace(" ", "&nbsp;")
        
        linha_celulas = [
            Paragraph(f"<nobr>{cod_formatado}</nobr>", estilo_td_codigo),
            Paragraph(str(desc), estilo_td_esq),
            Paragraph(f"<nobr>{est}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_comptar}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_divergencia}</nobr>", estilo_td_centro)
        ]
        
        if tem_na_comptar:
            contagens_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    val_tratado = str(int(float(locs_prod[i][1]))).replace('.', '&#46;')
                    contagens_vals.append(Paragraph(f"<nobr>{val_tratado}</nobr>", estilo_td_centro))
                else:
                    contagens_vals.append(Paragraph("", estilo_td_centro))
            
            locais_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    nome_loc_original = str(limpar_nome_local(locs_prod[i][0]))
                    nome_loc_seguro = nome_loc_original.replace('.', '&#46;').replace(' ', '&nbsp;')
                    locais_vals.append(Paragraph(f"<nobr>{nome_loc_seguro}</nobr>", estilo_td_centro))
                else:
                    locais_vals.append(Paragraph("", estilo_td_centro))
        else:
            contagens_vals = [Paragraph("<nobr>0</nobr>", estilo_td_centro)] + [Paragraph("", estilo_td_centro)] * (max_locais_encontrados - 1)
            locais_vals = [Paragraph("<nobr>Sem&nbsp;local</nobr>", estilo_td_centro)] + [Paragraph("", estilo_td_centro)] * (max_locais_encontrados - 1)
            
        linha_celulas.extend(contagens_vals)
        linha_celulas.extend(locais_vals)
        dados_tabela.append(linha_celulas)
        
    # 🛡️ Largura útil exata da página paisagem (Letter: 792 pt largura - 20 pt margens = 772 pt)
    largura_util = 772
    largura_codigo = 42
    largura_desc = 110
    largura_est = 22
    largura_comp = 22
    largura_div = 24
    largura_fixa_inicial = largura_codigo + largura_desc + largura_est + largura_comp + largura_div # 220 pt
    
    num_colunas = len(headers_pdf)
    colunas_restantes = num_colunas - 5
    # 🛡️ Distribui o espaço restante perfeitamente entre as colunas dinâmicas (C.1... e Loc.1...)
    largura_outras = max((largura_util - largura_fixa_inicial) / colunas_restantes, 24) if colunas_restantes > 0 else 24
    
    col_widths = [largura_codigo, largura_desc, largura_est, largura_comp, largura_div] + [largura_outras] * colunas_restantes
    
    tabela_pdf = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
    tabela_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('BACKGROUND', (4, 1), (4, -1), colors.HexColor('#D9D9D9')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
    ]))
    
    story.append(tabela_pdf)
    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

# ====================================================================================
# GERADOR DO EXCEL DE CONFERÊNCIA
# ====================================================================================
def gerar_planilha_filtrada_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Conferencia_{tipo_filtro}"
    ws.views.sheetView[0].showGridLines = True
    
    fonte_titulo = Font(name="Arial", size=11, bold=True, color="1F4E78")
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fonte_dados = Font(name="Arial", size=9)
    fonte_dados_negrito = Font(name="Arial", size=9, bold=True)
    
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_divergencia = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    borda_fina = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )
    
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    soma_comptar_por_produto = {}
    descricoes_comptar = {}
    
    for _, row in df_localizacao.iterrows():
        loc = limpar_nome_local(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, '')).strip().replace('.0', '')
        desc = str(row.get(col_desc_loc, '')).strip()
        try:
            qtd = int(float(row.get(col_qtd_loc, 0)))
        except Exception:
            qtd = 0
            
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
            soma_comptar_por_produto[cod] = 0
            descricoes_comptar[cod] = desc
        locais_por_produto[cod].append((loc, qtd))
        soma_comptar_por_produto[cod] += qtd
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    
    codigos_estoque_set = set(df_estoque[col_cod_est].astype(str).str.strip().str.replace('.0', '', regex=False).values)
    
    itens_filtrados_temp = []
    if tipo_filtro == "SEM_SALDO":
        for cod, locs in locais_por_produto.items():
            if cod not in codigos_estoque_set:
                desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
                soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
                itens_filtrados_temp.append((cod, desc_item, 0, True, soma_comptar_val, soma_comptar_val, locs))
    else:
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip().replace('.0', '')
            desc = str(row.get(col_desc_est, ''))
            try:
                est = int(float(row.get(col_est, 0)))
            except Exception:
                est = 0
            
            locs_prod = locais_por_produto.get(cod, [])
            tem_na_comptar = len(locs_prod) > 0
            soma_comptar_val = soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
            divergencia_val = soma_comptar_val - est
            
            incluir = False
            if tipo_filtro == "TODOS":
                incluir = True
            elif tipo_filtro == "ZERADOS":
                if not tem_na_comptar:
                    incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))
                
    itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1
    
    ws.cell(row=1, column=1, value=f"GRUPO KMB // RELATÓRIO DE CONFERÊNCIA ({tipo_filtro}) — LOJA: {numero_loja_input}").font = fonte_titulo
    
    headers_excel = ["Código", "Descrição", "EST.", "COMP.", "DIVERG."]
    for i in range(1, max_locais_encontrados + 1):
        headers_excel.append(f"C.{i}")
    for i in range(1, max_locais_encontrados + 1):
        headers_excel.append(f"Loc.{i}")
        
    for col_idx, h in enumerate(headers_excel, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
        
    row_idx = 4
    for cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod in itens_filtrados_temp:
        cel_comptar = soma_comptar_val if tem_na_comptar else 0
        cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)
        
        c_cod = ws.cell(row=row_idx, column=1, value=str(cod).strip())
        c_cod.font = fonte_dados_negrito
        c_cod.alignment = Alignment(horizontal="center", vertical="center")
        c_cod.border = borda_fina
        
        c_desc = ws.cell(row=row_idx, column=2, value=str(desc))
        c_desc.font = fonte_dados
        c_desc.alignment = Alignment(horizontal="left", vertical="center")
        c_desc.border = borda_fina
        
        c_est = ws.cell(row=row_idx, column=3, value=est)
        c_est.font = fonte_dados
        c_est.alignment = Alignment(horizontal="center", vertical="center")
        c_est.border = borda_fina
        
        c_comp = ws.cell(row=row_idx, column=4, value=cel_comptar)
        c_comp.font = fonte_dados
        c_comp.alignment = Alignment(horizontal="center", vertical="center")
        c_comp.border = borda_fina
        c_comp.fill = fill_divergencia
        
        c_div = ws.cell(row=row_idx, column=5, value=cel_divergencia)
        c_div.font = fonte_dados_negrito
        c_div.alignment = Alignment(horizontal="center", vertical="center")
        c_div.border = borda_fina
        
        for i in range(max_locais_encontrados):
            val_contagem = int(float(locs_prod[i][1])) if i < len(locs_prod) else ""
            c_c = ws.cell(row=row_idx, column=6 + i, value=val_contagem)
            c_c.font = fonte_dados
            c_c.alignment = Alignment(horizontal="center", vertical="center")
            c_c.border = borda_fina
            
        for i in range(max_locais_encontrados):
            nome_loc = limpar_nome_local(locs_prod[i][0]) if i < len(locs_prod) else ("Sem local" if i == 0 and not tem_na_comptar else "")
            c_l = ws.cell(row=row_idx, column=6 + max_locais_encontrados + i, value=nome_loc)
            c_l.font = fonte_dados
            c_l.alignment = Alignment(horizontal="center", vertical="center")
            c_l.border = borda_fina
            
        row_idx += 1
        
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter == 'B':
            ws.column_dimensions[col_letter].width = 40
        else:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
            
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ====================================================================================
# APLICAÇÃO STREAMLIT (MAIN)
# ====================================================================================
def main():
    st.set_page_config(page_title="GRUPO KMB // NOVA CORE 2", page_icon="⚡", layout="wide")
    SENHA_MESTRE = "novacore2026"
    
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        
    caminho_logo_tela = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
    
    if not st.session_state.autenticado:
        st.markdown("""
        <style>
        .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
        </style>
        """, unsafe_allow_html=True)
        
        col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
        with col_l2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            if caminho_logo_tela:
                st.image(caminho_logo_tela, use_container_width=True)
            else:
                st.warning("⚠️ Brasão não encontrado na pasta.")
            st.markdown("## ⚡ NOVA CORE 2 // AUTENTICAÇÃO")
            st.markdown("Insira a senha de acesso executivo para desbloquear o sistema.")
            
            senha_digitada = st.text_input("Senha de Acesso", type="password", key="input_senha")
            if st.button("🔓 Entrar no Sistema", use_container_width=True):
                if senha_digitada == SENHA_MESTRE:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta. Acesso negado.")
        return

    st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
    .main-header { padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        if caminho_logo_tela:
            st.image(caminho_logo_tela, use_container_width=True)
        st.markdown("### **GRUPO KMB**")
        st.markdown("*Nova Core 2 • Geração de Conferência*")
        st.markdown("---")
        if st.button("🔒 Encerrar Sessão (Sair)", use_container_width=True):
            st.session_state.autenticado = False
            st.rerun()
            
    st.markdown("""
    <div class="main-header">
        <div class="jarvis-title">⚡ GRUPO KMB // NOVA CORE 2</div>
        <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">MÓDULO EXCLUSIVO DE PLANILHA & PDF DE CONFERÊNCIA AUTOMATIZADA</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 GERAÇÃO AUTOMATIZADA DE LISTAS DE CONFERÊNCIA")
    st.markdown("Insira o número da loja e carregue os arquivos abaixo para gerar as listas com todas as colunas preservadas em linha única:")
    
    numero_loja_ui = st.number_input("Número da Loja (ex: 13016 ou 17315)", min_value=1, max_value=99999, value=13016, step=1)
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        arq_loc = st.file_uploader("1. Arquivo de Itens por Localização (ex: Itens_localizacao.csv)", type=['csv', 'xlsx', 'xls'], key="up_loc")
    with col_up2:
        arq_est = st.file_uploader("2. Arquivo de Saldo de Estoque (ex: saldo-estoque.csv)", type=['csv', 'xlsx', 'xls'], key="up_est")
        
    if arq_loc and arq_est:
        df_loc_temp, msg1 = processar_arquivo_upload(arq_loc)
        df_est_temp, msg2 = processar_arquivo_upload(arq_est)
        
        if df_loc_temp is not None and df_est_temp is not None:
            st.success("✅ Arquivos carregados e validados com sucesso!")
            st.markdown("---")
            
            st.markdown("### 👁️ Pré-visualização na Tela")
            buf_preview = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "TODOS")
            df_preview_exibir = pd.read_excel(buf_preview, sheet_name=0, skiprows=3, dtype=str).fillna("")
            
            cols_vistas = {}
            novas_colunas = []
            for c in df_preview_exibir.columns:
                c_str = str(c).strip()
                if '.' in c_str and c_str.split('.')[0].isdigit():
                    c_str = c_str.split('.')[0]
                if c_str in cols_vistas:
                    cols_vistas[c_str] += 1
                    novas_colunas.append(f"{c_str}_{cols_vistas[c_str]}")
                else:
                    cols_vistas[c_str] = 0
                    novas_colunas.append(c_str)
            df_preview_exibir.columns = novas_colunas
            
            for col in df_preview_exibir.columns:
                if "Descrição" not in col:
                    df_preview_exibir[col] = df_preview_exibir[col].apply(
                        lambda x: x[:-2] if isinstance(x, str) and x.endswith('.0') else x
                    )
                    
            st.dataframe(df_preview_exibir, use_container_width=True, height=420)
            
            st.download_button(
                label="📥 Baixar Pré-visualização em Excel (.xlsx)",
                data=buf_preview,
                file_name=f"Preview_SaldoLoja_Loja_{numero_loja_ui}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown("### 📥 1. Listas de Conferência em Excel (Ordem Alfabética)")
            
            c_imp1, c_imp2, c_imp3, c_imp4 = st.columns(4)
            with c_imp1:
                buf_zerados = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS")
                st.download_button("📥 Baixar Excel (Zerados)", data=buf_zerados, file_name=f"Excel_Zerados_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c_imp2:
                buf_faltas = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS")
                st.download_button("📥 Baixar Excel (Faltas)", data=buf_faltas, file_name=f"Excel_Faltas_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c_imp3:
                buf_sobras = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS")
                st.download_button("📥 Baixar Excel (Sobras)", data=buf_sobras, file_name=f"Excel_Sobras_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c_imp4:
                buf_sem_saldo = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO")
                st.download_button("📥 Baixar Excel (Sem Saldo)", data=buf_sem_saldo, file_name=f"Excel_SemSaldo_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
            data_atual_str = pd.Timestamp.now().strftime('%d.%m.%Y')
            st.markdown("---")
            st.markdown("### 📄 2. Listas de Conferência em PDF")
            
            c_pdf1, c_pdf2, c_pdf3, c_pdf4 = st.columns(4)
            with c_pdf1:
                pdf_zerados = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS")
                st.download_button("📑 Baixar PDF (Zerados)", data=pdf_zerados, file_name=f"{numero_loja_ui}.{data_atual_str}.zerados.pdf", mime="application/pdf", use_container_width=True)
            with c_pdf2:
                pdf_faltas = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS")
                st.download_button("📑 Baixar PDF (Faltas)", data=pdf_faltas, file_name=f"{numero_loja_ui}.{data_atual_str}.faltas.pdf", mime="application/pdf", use_container_width=True)
            with c_pdf3:
                pdf_sobras = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS")
                st.download_button("📑 Baixar PDF (Sobras)", data=pdf_sobras, file_name=f"{numero_loja_ui}.{data_atual_str}.sobras.pdf", mime="application/pdf", use_container_width=True)
            with c_pdf4:
                pdf_sem_saldo = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO")
                st.download_button("📑 Baixar PDF (Sem Saldo)", data=pdf_sem_saldo, file_name=f"{numero_loja_ui}.{data_atual_str}.sem_saldo.pdf", mime="application/pdf", use_container_width=True)
        else:
            st.error("❌ Erro ao processar os arquivos.")

if __name__ == "__main__":
    main()
