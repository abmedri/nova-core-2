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
