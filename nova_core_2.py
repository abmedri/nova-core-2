# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. - NOVA CORE 2 (VERSÃO COMPLETA SALVA)
# ====================================================================================
import io
import os
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import qrcode
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st


def limpar_colunas_duplicadas(df):
  novas_colunas = []
  contagem = {}
  for col in df.columns:
    col_str = str(col).strip()
    if col_str in contagem:
      contagem[col_str] += 1
      novas_colunas.append(f"{col_str}_{contagem[col_str]}")
    else:
      contagem[col_str] = 0
      novas_colunas.append(col_str)
  df.columns = novas_colunas
  return df


@st.cache_data
def processar_arquivo_upload(conteudo_bytes, nome_arquivo):
  nome = nome_arquivo.lower()
  try:
    arquivo_carregado = io.BytesIO(conteudo_bytes)
    if nome.endswith(".csv"):
      for enc in ["utf-8", "latin-1", "cp1252", "iso-8859-1"]:
        try:
          arquivo_carregado.seek(0)
          df = pd.read_csv(
              arquivo_carregado, encoding=enc, sep=";", low_memory=False
          )
          if df.shape[1] <= 1:
            arquivo_carregado.seek(0)
            df = pd.read_csv(
                arquivo_carregado, encoding=enc, sep=",", low_memory=False
            )
          df = limpar_colunas_duplicadas(df)
          col_mapping = {}
          for col in df.columns:
            col_lower = col.lower().strip()
            if "custo" in col_lower:
              col_mapping[col] = "Custo (R$)"
            elif "valor de ajuste" in col_lower:
              col_mapping[col] = "Valor de Ajuste (R$)"
            elif "valor de estoque" in col_lower:
              col_mapping[col] = "Valor de Estoque (R$)"
          df = df.rename(columns=col_mapping)
          for col in df.columns:
            if col.lower() not in [
                "código do produto",
                "codigo do produto",
                "loja",
                "inventario",
                "descrição do produto",
                "emb",
                "local",
                "data de abertura",
                "data de finalização",
                "tipo",
            ]:
              try:
                serie_str = (
                    df[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False)
                )
                conv_num = pd.to_numeric(serie_str, errors="coerce")
                if conv_num.notna().sum() > 0:
                  df[col] = conv_num
              except Exception:
                pass
          return df, "Inventário processado com sucesso."
        except Exception:
          continue
      return None, "Erro: Não foi possível decodificar o ficheiro CSV."
    elif nome.endswith((".xls", ".xlsx")):
      df = pd.read_excel(arquivo_carregado)
      df = limpar_colunas_duplicadas(df)
      return df, "Excel processado com sucesso."
    else:
      return None, "Formato não suportado."
  except Exception as e:
    return None, "Erro ao processar ficheiro: " + str(e)


def limpar_nome_local(val):
  s = str(val).strip()
  if s.endswith(".0"):
    s = s[:-2]
  return s


# ====================================================================================
# BLOCO 2: GERADOR DE ETIQUETAS QR CODE KMB
# ====================================================================================
def gerar_pdf_etiquetas_kmb(inicio, fim, caminho_brasao):
  buffer = io.BytesIO()
  c = canvas.Canvas(buffer, pagesize=A4)

  # Medidas exatas do modelo: 2 colunas x 7 linhas = 14 etiquetas por página A4
  margin_top = 38.0
  row_height = 104.0
  col_width = 272.0
  col_lefts = [20.0, 303.0]

  w_qr = 96.0
  w_num = 122.0
  w_brasao = 54.0

  brasao_reader = None
  if caminho_brasao and os.path.exists(caminho_brasao):
    try:
      brasao_reader = ImageReader(caminho_brasao)
    except Exception:
      brasao_reader = None

  total_etiquetas = (fim - inicio) + 1
  total_paginas = (total_etiquetas + 13) // 14

  def desenhar_rodape(pagina_atual):
    c.saveState()
    c.setFont("Helvetica", 8.5)
    c.setFillColorRGB(0.35, 0.35, 0.35)
    text = (
        f"Gerador de QrCode KMB  -  Página {pagina_atual} de {total_paginas}"
    )
    c.drawCentredString(A4[0] / 2.0, 18, text)
    c.restoreState()

  for idx, num in enumerate(range(inicio, fim + 1)):
    pos_na_pagina = idx % 14

    # Inicia nova página a cada 14 etiquetas (exceto na primeira)
    if idx > 0 and pos_na_pagina == 0:
      c.showPage()

    # Desenha o rodapé assim que inicia a página
    if pos_na_pagina == 0:
      pagina_atual = (idx // 14) + 1
      desenhar_rodape(pagina_atual)

    current_col = pos_na_pagina // 7
    current_row = pos_na_pagina % 7

    x = col_lefts[current_col]
    y = A4[1] - margin_top - ((current_row + 1) * row_height)

    x_qr = x
    x_num = x_qr + w_qr
    x_brasao = x_num + w_num

    # Bordas externas e divisórias internas
    c.setStrokeColorRGB(0.15, 0.15, 0.15)
    c.setLineWidth(0.8)
    c.rect(x, y, col_width, row_height)
    c.line(x_num, y, x_num, y + row_height)
    c.line(x_brasao, y, x_brasao, y + row_height)

    # QR Code em memória apenas com o número formatado (sem prefixo)
    str_num = f"{num:02d}"
    texto_qr = str_num
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(texto_qr)
    qr.make(fit=True)
    img_qr = qr.make_image(fill_color="black", back_color="white")

    qr_buf = io.BytesIO()
    img_qr.save(qr_buf, format="PNG")
    qr_buf.seek(0)
    c.drawImage(
        ImageReader(qr_buf),
        x_qr + 5,
        y + 5,
        width=w_qr - 10,
        height=row_height - 10,
        preserveAspectRatio=True,
    )

    # Número grande e encorpado preenchendo a célula central
    c.setFillColorRGB(0, 0, 0)
    if len(str_num) <= 2:
      font_size = 74
      y_offset = 26
    elif len(str_num) == 3:
      font_size = 54
      y_offset = 19
    else:
      font_size = 42
      y_offset = 15

    c.setFont("Helvetica-Bold", font_size)
    c.drawCentredString(
        x_num + (w_num / 2.0), y + (row_height / 2.0) - y_offset, str_num
    )

    # Brasão na terceira coluna
    if brasao_reader:
      c.drawImage(
          brasao_reader,
          x_brasao + 3,
          y + 8,
          width=w_brasao - 6,
          height=row_height - 16,
          preserveAspectRatio=True,
          mask="auto",
      )

  c.save()
  buffer.seek(0)
  return buffer


# ====================================================================================
# GERADOR DO EXCEL DE CONFERÊNCIA (COM BORDAS PRETAS NÍTIDAS)
# ====================================================================================
def gerar_planilha_filtrada_kmb(
    df_localizacao, df_estoque, numero_loja_input, tipo_filtro
):
  wb = openpyxl.Workbook()

  fonte_padrao = Font(name="Arial", size=10)
  fonte_descricao_excel = Font(name="Arial", size=9)
  fonte_negrito = Font(name="Arial", size=10, bold=True)
  fonte_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
  fill_cabecalho = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  fill_cinza_claro = PatternFill(
      start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
  )

  # Borda fina preta (#000000) para garantir forte destaque na impressão
  borda_fina = Border(
      left=Side(style="thin", color="000000"),
      right=Side(style="thin", color="000000"),
      top=Side(style="thin", color="000000"),
      bottom=Side(style="thin", color="000000"),
  )
  borda_destaque = Border(
      left=Side(style="medium", color="000000"),
      right=Side(style="medium", color="000000"),
      top=Side(style="medium", color="000000"),
      bottom=Side(style="medium", color="000000"),
  )

  col_loc = next(
      (c for c in df_localizacao.columns if "local" in c.lower()),
      df_localizacao.columns[2]
      if len(df_localizacao.columns) > 2
      else df_localizacao.columns[0],
  )
  col_cod_loc = next(
      (
          c
          for c in df_localizacao.columns
          if "codigo" in c.lower() or "cod" in c.lower()
      ),
      df_localizacao.columns[3]
      if len(df_localizacao.columns) > 3
      else df_localizacao.columns[0],
  )
  col_desc_loc = next(
      (c for c in df_localizacao.columns if "desc" in c.lower()),
      df_localizacao.columns[4]
      if len(df_localizacao.columns) > 4
      else df_localizacao.columns[0],
  )
  col_qtd_loc = next(
      (
          c
          for c in df_localizacao.columns
          if "qtd" in c.lower() or "quantidade" in c.lower()
      ),
      df_localizacao.columns[5]
      if len(df_localizacao.columns) > 5
      else df_localizacao.columns[0],
  )

  locais_por_produto = {}
  soma_comptar_por_produto = {}
  descricoes_comptar = {}

  for _, row in df_localizacao.iterrows():
    loc = limpar_nome_local(row.get(col_loc, ""))
    cod = str(row.get(col_cod_loc, "")).strip()
    desc = str(row.get(col_desc_loc, "")).strip()
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

  col_cod_est = next(
      (c for c in df_estoque.columns if "codigo" in c.lower()), "Codigo"
  )
  col_desc_est = next(
      (c for c in df_estoque.columns if "desc" in c.lower()), "Descrição"
  )
  col_est = next(
      (c for c in df_estoque.columns if "estoque" in c.lower()), "Estoque"
  )

  codigos_estoque_set = (
      set(df_estoque[col_cod_est].astype(str).str.strip().values)
      if col_cod_est in df_estoque.columns
      else set()
  )
  itens_filtrados_temp = []

  if tipo_filtro == "SEM_SALDO":
    for cod, locs in locais_por_produto.items():
      if cod not in codigos_estoque_set:
        desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
        soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
        divergencia_val = soma_comptar_val - 0
        itens_filtrados_temp.append((
            cod,
            desc_item,
            0,
            True,
            soma_comptar_val,
            divergencia_val,
            locs,
        ))
  elif tipo_filtro == "ZERADOS":
    for idx, row in df_estoque.iterrows():
      cod = str(row.get(col_cod_est, "")).strip()
      if cod not in locais_por_produto:
        desc = row.get(col_desc_est, "")
        try:
          est = int(float(row.get(col_est, 0)))
        except Exception:
          est = 0
        itens_filtrados_temp.append((cod, desc, est, False, 0, -est, []))
  else:
    for idx, row in df_estoque.iterrows():
      cod = str(row.get(col_cod_est, "")).strip()
      desc = str(row.get(col_desc_est, ""))
      try:
        est = int(float(row.get(col_est, 0)))
      except Exception:
        est = 0
      locs_prod = locais_por_produto.get(cod, [])
      tem_na_comptar = len(locs_prod) > 0
      soma_comptar_val = (
          soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
      )
      divergencia_val = soma_comptar_val - est

      incluir = False
      if tipo_filtro == "TODOS":
        incluir = True
      elif tipo_filtro == "FALTAS":
        if tem_na_comptar and divergencia_val < 0:
          incluir = True
      elif tipo_filtro == "SOBRAS":
        if tem_na_comptar and divergencia_val > 0:
          incluir = True

      if incluir:
        itens_filtrados_temp.append((
            cod,
            desc,
            est,
            tem_na_comptar,
            soma_comptar_val,
            divergencia_val,
            locs_prod,
        ))

  itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
  locais_com_contagem = [
      len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0
  ]
  max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1

  nome_aba = (
      "Itens Sem Saldo"
      if tipo_filtro == "SEM_SALDO"
      else ("Zerados" if tipo_filtro == "ZERADOS" else "Saldo Loja")
  )
  ws_saldo = wb.active
  ws_saldo.title = nome_aba
  ws_saldo.views.sheetView[0].showGridLines = True

  ws_saldo.page_setup.orientation = ws_saldo.ORIENTATION_LANDSCAPE
  ws_saldo.page_setup.paperSize = ws_saldo.PAPERSIZE_A4
  ws_saldo.sheet_properties.pageSetUpPr.fitToPage = True
  ws_saldo.page_setup.fitToWidth = 1
  ws_saldo.page_setup.fitToHeight = 0

  ws_saldo.page_margins.left = 0.2
  ws_saldo.page_margins.right = 0.2
  ws_saldo.page_margins.top = 0.3
  ws_saldo.page_margins.bottom = 0.3
  ws_saldo.page_margins.header = 0.1
  ws_saldo.page_margins.footer = 0.1

  ws_saldo.oddFooter.left.text = f"Loja: {numero_loja_input} | Grupo KMB"
  ws_saldo.oddFooter.center.text = "Exatidão • Controle • Eficiência"
  ws_saldo.oddFooter.right.text = "Página &P de &N"

  caminho_brasao_timbre = (
      "brasao_sem_fundo.png"
      if os.path.exists("brasao_sem_fundo.png")
      else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
  )
  if caminho_brasao_timbre:
    try:
      img_timbre = OpenpyxlImage(caminho_brasao_timbre)
      img_timbre.width = 110
      img_timbre.height = 42
      ws_saldo.add_image(img_timbre, "A1")
    except Exception:
      pass

  ws_saldo["D1"] = "DATA:"
  ws_saldo["D1"].font = fonte_negrito
  ws_saldo["D1"].alignment = Alignment(horizontal="center", vertical="center")
  ws_saldo["D1"].border = borda_destaque

  ws_saldo["E1"] = pd.Timestamp.now().strftime("%d/%m/%Y")
  ws_saldo["E1"].font = fonte_padrao
  ws_saldo["E1"].alignment = Alignment(horizontal="center", vertical="center")
  ws_saldo["E1"].border = borda_destaque

  ws_saldo["D2"] = "LOJA:"
  ws_saldo["D2"].font = fonte_negrito
  ws_saldo["D2"].alignment = Alignment(horizontal="center", vertical="center")
  ws_saldo["D2"].border = borda_destaque

  ws_saldo["E2"] = int(numero_loja_input)
  ws_saldo["E2"].font = fonte_negrito
  ws_saldo["E2"].alignment = Alignment(horizontal="center", vertical="center")
  ws_saldo["E2"].border = borda_destaque

  headers_saldo = ["Codigo", "Descrição", "Estoque", "COMPTAR", "DIVERGÊNCIA"]
  for i in range(1, max_locais_encontrados + 1):
    if i == 1:
      headers_saldo.append("CONT. 1")
    else:
      headers_saldo.append(str(int(i)))
  for i in range(1, max_locais_encontrados + 1):
    if i == 1:
      headers_saldo.append("Local 1")
    else:
      headers_saldo.append(str(int(i)))

  ws_saldo.row_dimensions[1].height = 24
  ws_saldo.row_dimensions[2].height = 20
  ws_saldo.row_dimensions[3].height = 10
  ws_saldo.row_dimensions[4].height = 22

  for col_idx, h_text in enumerate(headers_saldo, start=1):
    cell = ws_saldo.cell(
        row=4, column=col_idx, value=str(h_text).replace(".0", "")
    )
    cell.font = fonte_cabecalho
    cell.fill = fill_cabecalho
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = borda_fina
  ws_saldo.print_title_rows = "4:4"
  ws_saldo.auto_filter.ref = (
      f"A4:{get_column_letter(len(headers_saldo))}{ws_saldo.max_row}"
  )

  for (
      cod,
      desc,
      est,
      tem_na_comptar,
      soma_comptar_val,
      divergencia_val,
      locs_prod,
  ) in itens_filtrados_temp:
    cel_comptar = soma_comptar_val if tem_na_comptar else 0
    cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)

    linha_dados = [cod, desc, est, cel_comptar, cel_divergencia]

    if tem_na_comptar:
      contagens_vals = []
      for i in range(max_locais_encontrados):
        if i < len(locs_prod):
          val_num = locs_prod[i][1]
          try:
            val_int = int(float(val_num))
            contagens_vals.append(val_int)
          except Exception:
            contagens_vals.append(val_num)
        else:
          contagens_vals.append("")

      locais_vals = [
          limpar_nome_local(locs_prod[i][0]) if i < len(locs_prod) else ""
          for i in range(max_locais_encontrados)
      ]
    else:
      if tipo_filtro == "ZERADOS":
        contagens_vals = ["0 - Comptar"] + [""] * (max_locais_encontrados - 1)
      else:
        contagens_vals = ["0"] + [""] * (max_locais_encontrados - 1)
      locais_vals = ["Sem local"] + [""] * (max_locais_encontrados - 1)

    linha_dados.extend(contagens_vals)
    linha_dados.extend(locais_vals)
    ws_saldo.append(linha_dados)

  for row in ws_saldo.iter_rows(
      min_row=5, max_row=ws_saldo.max_row, min_col=1, max_col=len(headers_saldo)
  ):
    for cell in row:
      if cell.column == 2:
        cell.font = fonte_descricao_excel
      else:
        cell.font = fonte_padrao

      cell.border = borda_fina
      if cell.row >= 5 and cell.column >= 3:
        if isinstance(cell.value, (int, float)):
          cell.number_format = "0"

      if cell.column == 5:
        cell.fill = fill_cinza_claro
        cell.font = Font(name="Arial", size=10, bold=True)

      if cell.column in [1, 3, 4, 5]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      elif cell.column == 2:
        cell.alignment = Alignment(horizontal="left", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="center", vertical="center")

  for col in ws_saldo.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
      if cell.row >= 4 and cell.value is not None:
        val_str = str(cell.value)
        if not val_str.startswith("="):
          max_len = max(max_len, len(val_str))
    ws_saldo.column_dimensions[col_letter].width = max(max_len + 2, 7)

  buffer = io.BytesIO()
  wb.save(buffer)
  buffer.seek(0)
  return buffer


# ====================================================================================
# GERADOR DE PDF DE CONFERÊNCIA OTIMIZADO (MULTIPLICADOR CORRIGIDO PARA 4.2)
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
    self.setFillColor(colors.HexColor("#64748b"))

    largura_pagina, altura_pagina = self._pagesize
    margem = 10

    self.setStrokeColor(colors.HexColor("#000000"))
    self.setLineWidth(0.5)
    self.line(margem, 22, largura_pagina - margem, 22)

    caminho_brasao = (
        "brasao_sem_fundo.png"
        if os.path.exists("brasao_sem_fundo.png")
        else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
    )
    if caminho_brasao:
      try:
        self.drawImage(
            caminho_brasao,
            margem,
            6,
            width=45,
            height=13,
            preserveAspectRatio=True,
            mask="auto",
        )
      except Exception:
        pass

    texto_esquerda = (
        "GRUPO KMB // Auditoria Patrimonial • Exatidão • Controle • Eficiência"
    )
    self.drawString(margem + 50, 10, texto_esquerda)

    texto_pagina = f"Página {self._pageNumber} de {page_count}"
    self.drawRightString(largura_pagina - margem, 10, texto_pagina)
    self.restoreState()


def gerar_pdf_conferencia_kmb(
    df_localizacao, df_estoque, numero_loja_input, tipo_filtro
):
  col_loc = next(
      (c for c in df_localizacao.columns if "local" in c.lower()),
      df_localizacao.columns[2]
      if len(df_localizacao.columns) > 2
      else df_localizacao.columns[0],
  )
  col_cod_loc = next(
      (
          c
          for c in df_localizacao.columns
          if "codigo" in c.lower() or "cod" in c.lower()
      ),
      df_localizacao.columns[3]
      if len(df_localizacao.columns) > 3
      else df_localizacao.columns[0],
  )
  col_desc_loc = next(
      (c for c in df_localizacao.columns if "desc" in c.lower()),
      df_localizacao.columns[4]
      if len(df_localizacao.columns) > 4
      else df_localizacao.columns[0],
  )
  col_qtd_loc = next(
      (
          c
          for c in df_localizacao.columns
          if "qtd" in c.lower() or "quantidade" in c.lower()
      ),
      df_localizacao.columns[5]
      if len(df_localizacao.columns) > 5
      else df_localizacao.columns[0],
  )

  locais_por_produto = {}
  soma_comptar_por_produto = {}
  descricoes_comptar = {}

  for _, row in df_localizacao.iterrows():
    loc = limpar_nome_local(row.get(col_loc, ""))
    cod = str(row.get(col_cod_loc, "")).strip()
    desc = str(row.get(col_desc_loc, "")).strip()
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

  col_cod_est = next(
      (c for c in df_estoque.columns if "codigo" in c.lower()), "Codigo"
  )
  col_desc_est = next(
      (c for c in df_estoque.columns if "desc" in c.lower()), "Descrição"
  )
  col_est = next(
      (c for c in df_estoque.columns if "estoque" in c.lower()), "Estoque"
  )

  codigos_estoque_set = (
      set(df_estoque[col_cod_est].astype(str).str.strip().values)
      if col_cod_est in df_estoque.columns
      else set()
  )

  itens_filtrados_temp = []
  if tipo_filtro == "SEM_SALDO":
    for cod, locs in locais_por_produto.items():
      if cod not in codigos_estoque_set:
        desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
        soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
        divergencia_val = soma_comptar_val - 0
        itens_filtrados_temp.append((
            cod,
            desc_item,
            0,
            True,
            soma_comptar_val,
            divergencia_val,
            locs,
        ))
  elif tipo_filtro == "ZERADOS":
    for idx, row in df_estoque.iterrows():
      cod = str(row.get(col_cod_est, "")).strip()
      if cod not in locais_por_produto:
        desc = str(row.get(col_desc_est, ""))
        try:
          est = int(float(row.get(col_est, 0)))
        except Exception:
          est = 0
        itens_filtrados_temp.append((cod, desc, est, False, 0, -est, []))
  else:
    for idx, row in df_estoque.iterrows():
      cod = str(row.get(col_cod_est, "")).strip()
      desc = str(row.get(col_desc_est, ""))
      try:
        est = int(float(row.get(col_est, 0)))
      except Exception:
        est = 0
      locs_prod = locais_por_produto.get(cod, [])
      tem_na_comptar = len(locs_prod) > 0
      soma_comptar_val = (
          soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
      )
      divergencia_val = soma_comptar_val - est

      incluir = False
      if tipo_filtro == "TODOS":
        incluir = True
      elif tipo_filtro == "FALTAS":
        if tem_na_comptar and divergencia_val < 0:
          incluir = True
      elif tipo_filtro == "SOBRAS":
        if tem_na_comptar and divergencia_val > 0:
          incluir = True

      if incluir:
        itens_filtrados_temp.append((
            cod,
            desc,
            est,
            tem_na_comptar,
            soma_comptar_val,
            divergencia_val,
            locs_prod,
        ))

  itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
  locais_com_contagem = [
      len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0
  ]
  max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1

  buffer = io.BytesIO()

  doc = SimpleDocTemplate(
      buffer,
      pagesize=landscape(letter),
      rightMargin=14,
      leftMargin=14,
      topMargin=18,
      bottomMargin=28,
  )

  styles = getSampleStyleSheet()
  estilo_titulo = ParagraphStyle(
      "TituloPDF",
      parent=styles["Heading1"],
      fontName="Helvetica-Bold",
      fontSize=9,
      textColor=colors.HexColor("#1F4E78"),
  )
  estilo_th = ParagraphStyle(
      "TH_Auto",
      fontName="Helvetica-Bold",
      fontSize=5.0,
      leading=6.0,
      textColor=colors.white,
      alignment=1,
  )

  estilo_td_codigo = ParagraphStyle(
      "TD_Codigo",
      fontName="Helvetica-Bold",
      fontSize=5.5,
      leading=6.5,
      textColor=colors.HexColor("#111827"),
      alignment=1,
  )
  estilo_td_esq = ParagraphStyle(
      "TD_Esq",
      fontName="Helvetica",
      fontSize=5.5,
      leading=6.5,
      textColor=colors.HexColor("#111827"),
      alignment=0,
      leftIndent=0,
      firstLineIndent=0,
      rightIndent=0,
  )
  estilo_td_negrito_centro = ParagraphStyle(
      "TD_Negrito_Centro",
      fontName="Helvetica-Bold",
      fontSize=6.0,
      leading=7.0,
      textColor=colors.HexColor("#111827"),
      alignment=1,
  )

  story = []
  story.append(
      Paragraph(
          f"<b>GRUPO KMB // RELATÓRIO DE CONFERÊNCIA ({tipo_filtro})</b> — LOJA:"
          f" {numero_loja_input} | DATA:"
          f" {pd.Timestamp.now().strftime('%d/%m/%Y')}",
          estilo_titulo,
      )
  )
  story.append(Spacer(1, 3))

  headers_pdf = ["Código", "Descrição", "EST.", "COMP.", "DIVERG."]
  for i in range(1, max_locais_encontrados + 1):
    headers_pdf.append(f"C.{int(i)}" if i == 1 else str(int(i)))
  for i in range(1, max_locais_encontrados + 1):
    headers_pdf.append(f"Loc.{int(i)}" if i == 1 else str(int(i)))

  dados_tabela = [[
      Paragraph(
          f"<nobr>{str(h).strip().replace('.0', '').replace(' ', '&nbsp;')}</nobr>",
          estilo_th,
      )
      for h in headers_pdf
  ]]

  max_chars_por_coluna = [0] * (max_locais_encontrados * 2)

  for (
      cod,
      desc,
      est,
      tem_na_comptar,
      soma_comptar_val,
      divergencia_val,
      locs_prod,
  ) in itens_filtrados_temp:
    cel_comptar = soma_comptar_val if tem_na_comptar else 0
    cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)

    cod_formatado = str(cod).strip().replace(" ", "&nbsp;")
    desc_limpa = str(desc).strip().replace(" ", "&nbsp;")

    linha_celulas = [
        Paragraph(f"<nobr>{cod_formatado}</nobr>", estilo_td_codigo),
        Paragraph(f"<nobr>{desc_limpa}</nobr>", estilo_td_esq),
        Paragraph(f"<nobr>{est}</nobr>", estilo_td_negrito_centro),
        Paragraph(f"<nobr>{cel_comptar}</nobr>", estilo_td_negrito_centro),
        Paragraph(f"<nobr>{cel_divergencia}</nobr>", estilo_td_negrito_centro),
    ]

    contagens_vals = []
    locais_vals = []
    if tem_na_comptar:
      for i in range(max_locais_encontrados):
        if i < len(locs_prod):
          val_qtd_str = str(int(float(locs_prod[i][1])))
          contagens_vals.append(
              Paragraph(f"<nobr>{val_qtd_str}</nobr>", estilo_td_negrito_centro)
          )
          max_chars_por_coluna[i] = max(
              max_chars_por_coluna[i], len(val_qtd_str)
          )

          nome_loc_original = str(limpar_nome_local(locs_prod[i][0]))
          nome_loc_seguro = nome_loc_original.replace(" ", "&nbsp;")
          locais_vals.append(
              Paragraph(
                  f"<nobr>{nome_loc_seguro}</nobr>", estilo_td_negrito_centro
              )
          )
          max_chars_por_coluna[max_locais_encontrados + i] = max(
              max_chars_por_coluna[max_locais_encontrados + i],
              len(nome_loc_original),
          )
        else:
          contagens_vals.append(Paragraph("", estilo_td_negrito_centro))
          locais_vals.append(Paragraph("", estilo_td_negrito_centro))
    else:
      if tipo_filtro == "ZERADOS":
        contagens_vals = [
            Paragraph("<nobr>0&nbsp;-&nbsp;Comptar</nobr>", estilo_td_negrito_centro)
        ] + [Paragraph("", estilo_td_negrito_centro)] * (
            max_locais_encontrados - 1
        )
        locais_vals = [
            Paragraph("<nobr>Sem&nbsp;local</nobr>", estilo_td_negrito_centro)
        ] + [Paragraph("", estilo_td_negrito_centro)] * (
            max_locais_encontrados - 1
        )
        max_chars_por_coluna[0] = max(max_chars_por_coluna[0], 12)
        max_chars_por_coluna[max_locais_encontrados] = max(
            max_chars_por_coluna[max_locais_encontrados], 9
        )
      else:
        contagens_vals = [
            Paragraph("<nobr>0</nobr>", estilo_td_negrito_centro)
        ] + [Paragraph("", estilo_td_negrito_centro)] * (
            max_locais_encontrados - 1
        )
        locais_vals = [
            Paragraph("<nobr>Sem&nbsp;local</nobr>", estilo_td_negrito_centro)
        ] + [Paragraph("", estilo_td_negrito_centro)] * (
            max_locais_encontrados - 1
        )
        max_chars_por_coluna[0] = max(max_chars_por_coluna[0], 1)
        max_chars_por_coluna[max_locais_encontrados] = max(
            max_chars_por_coluna[max_locais_encontrados], 9
        )

    linha_celulas.extend(contagens_vals)
    linha_celulas.extend(locais_vals)
    dados_tabela.append(linha_celulas)

  largura_util = 792 - 28

  largura_codigo = 45
  max_len_desc = (
      max([len(str(item[1]).strip()) for item in itens_filtrados_temp])
      if itens_filtrados_temp
      else 20
  )
  largura_desc = max(110.0, min(float(max_len_desc * 4.2), 250.0))

  largura_est = 28
  largura_comp = 28
  largura_div = 30

  largura_fixa_inicial = (
      largura_codigo + largura_desc + largura_est + largura_comp + largura_div
  )
  largura_restante_total = largura_util - largura_fixa_inicial

  pesos_colunas = []
  for idx_col, max_c in enumerate(max_chars_por_coluna):
    header_text = (
        headers_pdf[5 + idx_col] if (5 + idx_col) < len(headers_pdf) else ""
    )
    peso = max(max_c, len(header_text), 3)
    pesos_colunas.append(peso)

  soma_pesos = sum(pesos_colunas) if sum(pesos_colunas) > 0 else 1
  col_widths = [
      largura_codigo,
      largura_desc,
      largura_est,
      largura_comp,
      largura_div,
  ]

  for peso in pesos_colunas:
    L_calc = (peso / soma_pesos) * largura_restante_total
    col_widths.append(max(L_calc, 18.0))

  tabela_pdf = Table(dados_tabela, colWidths=col_widths, repeatRows=1)

  idx_col_divergencia = 4
  idx_col_c1 = 5
  idx_col_loc1 = 5 + max_locais_encontrados

  tabela_pdf.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
          (
              "BACKGROUND",
              (idx_col_divergencia, 1),
              (idx_col_divergencia, -1),
              colors.HexColor("#D9D9D9"),
          ),
          (
              "BACKGROUND",
              (idx_col_c1, 1),
              (idx_col_c1, -1),
              colors.HexColor("#F1F5F9"),
          ),
          (
              "BACKGROUND",
              (idx_col_loc1, 1),
              (idx_col_loc1, -1),
              colors.HexColor("#F1F5F9"),
          ),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 1.0),
          ("TOPPADDING", (0, 0), (-1, -1), 1.0),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#000000")),
      ])
  )

  story.append(tabela_pdf)
  doc.build(story, canvasmaker=NumberedCanvas)
  buffer.seek(0)
  return buffer


# ====================================================================================
# BLOCO 3: AUTOMAÇÃO DE RELATÓRIO DE INVENTÁRIO (MEMÓRIA, JUSTIFICAR & RESUMO)
# ====================================================================================
def _b5_ler_csv_inventario(uploaded_file):
  """Lê o arquivo CSV exportado do sistema tratando codificação, separador e colunas monetárias."""
  raw_bytes = uploaded_file.getvalue()
  for enc in ["latin1", "cp1252", "utf-8-sig", "utf-8"]:
    try:
      df = pd.read_csv(io.BytesIO(raw_bytes), sep=";", encoding=enc)
      if "Código do Produto" in df.columns:
        for col in [
            "Custo (R$)",
            "Valor de Ajuste (R$)",
            "Valor de Estoque (R$)",
        ]:
          if col in df.columns:
            df[col] = (
                pd.to_numeric(
                    df[col]
                    .astype(str)
                    .str.replace(".", "", regex=False)
                    .str.replace(",", ".", regex=False),
                    errors="coerce",
                )
                .fillna(0.0)
            )
        return df
    except Exception:
      continue
  raise ValueError(
      "Não foi possível ler o CSV. Verifique se o separador é ';' e se a"
      " coluna 'Código do Produto' existe."
  )


def _b5_ler_inventario_anterior(uploaded_file):
  """Lê o arquivo CSV (ou Excel) do inventário anterior e retorna um dicionário {Código: Ajuste}."""
  if uploaded_file is None:
    return {}
  raw_bytes = uploaded_file.getvalue()
  nome_arq = str(getattr(uploaded_file, "name", "")).lower()

  if nome_arq.endswith(".csv"):
    df_ant = None
    for enc in ["latin1", "cp1252", "utf-8-sig", "utf-8"]:
      for sep in [";", ","]:
        try:
          temp = pd.read_csv(io.BytesIO(raw_bytes), sep=sep, encoding=enc)
          if (
              "Código do Produto" in temp.columns
              and "Ajuste" in temp.columns
          ):
            df_ant = temp
            break
        except Exception:
          continue
      if df_ant is not None:
        break
    if df_ant is None:
      raise ValueError(
          "O arquivo CSV do inventário anterior precisa ter as colunas"
          " 'Código do Produto' e 'Ajuste' (separador ';' ou ',')."
      )
  else:
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    df_ant = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=xls.sheet_names[0])
    if (
        "Código do Produto" not in df_ant.columns
        or "Ajuste" not in df_ant.columns
    ):
      raise ValueError(
          "O arquivo do inventário anterior precisa ter as colunas 'Código do"
          " Produto' e 'Ajuste' na primeira aba."
      )

  df_ant = df_ant.drop_duplicates(subset=["Código do Produto"], keep="last")
  return df_ant.set_index("Código do Produto")["Ajuste"].to_dict()


def _b5_gerar_excel_memoria(df_atual, mapa_anterior, nome_arquivo_csv):
  nome_aba_base = os.path.splitext(nome_arquivo_csv)[0][:31]

  # Regra da aba JUSTIFICAR:
  # Negativos (< 0), Sobras (> 0) ou Congelado == 0
  condicao = (
      (df_atual["Ajuste"] < 0)
      | (df_atual["Ajuste"] > 0)
      | (df_atual["Congelado"] == 0)
  )
  df_just = df_atual[condicao].copy()
  df_just = df_just.sort_values(
      by=["Ajuste", "Descrição do Produto"], ascending=[True, True]
  ).reset_index(drop=True)

  wb = openpyxl.Workbook()

  # Paleta de estilos
  navy_dark = "1E3A5F"
  zebra_fill_hex = "F7FAFC"
  border_gray = "CBD5E0"
  total_fill_hex = "E2E8F0"
  ant_header_fill = "FEF3C7"
  ant_header_font = "92400E"
  ant_cell_fill = "FFFBEB"

  font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  fill_header = PatternFill(
      fill_type="solid", start_color=navy_dark, end_color=navy_dark
  )
  font_data = Font(name="Calibri", size=11, color="1A202C")
  font_bold = Font(name="Calibri", size=11, bold=True, color="1A202C")
  fill_zebra = PatternFill(
      fill_type="solid", start_color=zebra_fill_hex, end_color=zebra_fill_hex
  )
  fill_total = PatternFill(
      fill_type="solid", start_color=total_fill_hex, end_color=total_fill_hex
  )

  thin_border = Border(
      left=Side(style="thin", color=border_gray),
      right=Side(style="thin", color=border_gray),
      top=Side(style="thin", color=border_gray),
      bottom=Side(style="thin", color=border_gray),
  )
  total_border = Border(
      top=Side(style="medium", color=navy_dark),
      bottom=Side(style="medium", color=navy_dark),
      left=Side(style="thin", color=border_gray),
      right=Side(style="thin", color=border_gray),
  )

  # =========================================================================
  # ABA 1: CÓPIA DO ARQUIVO CSV
  # =========================================================================
  ws_base = wb.active
  ws_base.title = nome_aba_base

  headers_base = list(df_atual.columns)
  ws_base.append(headers_base)
  for col_idx in range(1, len(headers_base) + 1):
    c = ws_base.cell(1, col_idx)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    c.border = thin_border
  ws_base.row_dimensions[1].height = 26

  for row_data in df_atual.itertuples(index=False):
    ws_base.append(list(row_data))

  for r_idx in range(2, ws_base.max_row + 1):
    for c_idx in range(1, ws_base.max_column + 1):
      c = ws_base.cell(r_idx, c_idx)
      c.font = font_data
      c.border = thin_border
      if r_idx % 2 == 1:
        c.fill = fill_zebra
      if c_idx in [1, 3, 4, 5, 6, 7, 13]:
        c.alignment = Alignment(horizontal="center", vertical="center")
      elif c_idx in [8, 9, 10]:
        c.number_format = "R$ #,##0.00"
        c.alignment = Alignment(horizontal="right", vertical="center")

  ws_base.freeze_panes = "A2"
  ws_base.auto_filter.ref = (
      f"A1:{get_column_letter(ws_base.max_column)}{ws_base.max_row}"
  )

  larguras_base = [
      18,
      42,
      8,
      12,
      11,
      10,
      18,
      14,
      19,
      19,
      36,
      12,
      13,
      18,
      22,
      13,
  ]
  for i, w in enumerate(larguras_base[: ws_base.max_column], start=1):
    ws_base.column_dimensions[get_column_letter(i)].width = w

  # =========================================================================
  # ABA 2: JUSTIFICAR
  # =========================================================================
  ws_just = wb.create_sheet(title="JUSTIFICAR")

  headers_just = [
      "Código do Produto",
      "Descrição do Produto",
      "Emb",
      "Congelado",
      "Digitado",
      "Ajuste",
      "Saldo em Estoque",
      "Ajuste Inventario Anterior",
      "JUSTIFICAR",
      "LOGISTICA",
  ]
  ws_just.append(headers_just)

  for col_idx in range(1, len(headers_just) + 1):
    c = ws_just.cell(1, col_idx)
    if col_idx == 8:
      c.fill = PatternFill(
          fill_type="solid", start_color=ant_header_fill, end_color=ant_header_fill
      )
      c.font = Font(name="Calibri", size=11, bold=True, color=ant_header_font)
    else:
      c.fill = fill_header
      c.font = font_header
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    c.border = thin_border
  ws_just.row_dimensions[1].height = 28

  max_base_row = ws_base.max_row
  sheet_ref = f"'{nome_aba_base}'"

  for idx, row in df_just.iterrows():
    r_excel = idx + 2
    cod = int(row["Código do Produto"])
    ajuste_ant = mapa_anterior.get(cod, 0)
    ajuste_ant = 0 if pd.isna(ajuste_ant) else int(ajuste_ant)

    row_vals = [
        cod,
        (
            f"=INDEX({sheet_ref}!$B$2:$B${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        (
            f"=INDEX({sheet_ref}!$C$2:$C${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        (
            f"=INDEX({sheet_ref}!$D$2:$D${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        (
            f"=INDEX({sheet_ref}!$E$2:$E${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        (
            f"=INDEX({sheet_ref}!$F$2:$F${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        (
            f"=INDEX({sheet_ref}!$G$2:$G${max_base_row},"
            f" MATCH(A{r_excel}, {sheet_ref}!$A$2:$A${max_base_row}, 0))"
        ),
        ajuste_ant,
        None,
        None,
    ]
    ws_just.append(row_vals)

    for c_idx in range(1, 11):
      cell = ws_just.cell(r_excel, c_idx)
      cell.border = thin_border
      if c_idx == 8:
        cell.font = Font(
            name="Calibri", size=11, bold=True, color=ant_header_font
        )
        cell.fill = PatternFill(
            fill_type="solid", start_color=ant_cell_fill, end_color=ant_cell_fill
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.font = font_data
        if idx % 2 == 1:
          cell.fill = fill_zebra
        if c_idx in [1, 3, 4, 5, 6, 7]:
          cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
          cell.alignment = Alignment(horizontal="left", vertical="center")

  total_row = ws_just.max_row + 1
  ws_just.cell(total_row, 1, "TOTAL").font = font_bold
  ws_just.cell(total_row, 1).alignment = Alignment(
      horizontal="center", vertical="center"
  )
  ws_just.cell(
      total_row, 2, f'=COUNTA(A2:A{total_row-1}) & " itens"'
  ).font = font_bold

  for c_idx in [4, 5, 6, 7, 8]:
    col_let = get_column_letter(c_idx)
    c = ws_just.cell(
        total_row, c_idx, f"=SUM({col_let}2:{col_let}{total_row-1})"
    )
    c.font = font_bold
    c.alignment = Alignment(horizontal="center", vertical="center")

  for c_idx in range(1, 11):
    cell = ws_just.cell(total_row, c_idx)
    cell.fill = fill_total
    cell.border = total_border

  red_fill = PatternFill(
      start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"
  )
  red_font = Font(color="991B1B", bold=True)
  green_fill = PatternFill(
      start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
  )
  green_font = Font(color="166534", bold=True)

  ws_just.conditional_formatting.add(
      f"F2:F{total_row-1}",
      CellIsRule(
          operator="lessThan", formula=["0"], fill=red_fill, font=red_font
      ),
  )
  ws_just.conditional_formatting.add(
      f"F2:F{total_row-1}",
      CellIsRule(
          operator="greaterThan",
          formula=["0"],
          fill=green_fill,
          font=green_font,
      ),
  )

  dv_just = DataValidation(
      type="list",
      formula1=(
          '"Furto / Perda,Erro de Contagem / Bipagem,Avaria /'
          " Vencido,Divergência de Recebimento,Troca / Devolução não"
          ' lançada,Produto sem saldo (Congelado 0),Outros"'
      ),
      allow_blank=True,
  )
  ws_just.add_data_validation(dv_just)
  dv_just.add(f"I2:I{total_row-1}")

  larguras_just = [18, 44, 8, 12, 11, 10, 18, 25, 32, 30]
  for i, w in enumerate(larguras_just, start=1):
    ws_just.column_dimensions[get_column_letter(i)].width = w

  ws_just.freeze_panes = "A2"
  ws_just.auto_filter.ref = f"A1:J{total_row-1}"

  # =========================================================================
  # ABA 3: RESUMO (Valores calculados em Python para funcionar em PC e Celular)
  # =========================================================================
  ws_res = wb.create_sheet(title="RESUMO")

  loja_titulo = (
      str(df_atual["Loja"].iloc[0])
      if "Loja" in df_atual.columns and len(df_atual) > 0
      else nome_aba_base
  )
  ws_res["A1"] = f"RESUMO DO INVENTÁRIO — {loja_titulo}"
  ws_res["A1"].font = Font(
      name="Calibri", size=13, bold=True, color=navy_dark
  )
  ws_res.merge_cells("A1:D1")

  res_headers = ["Indicador", "Qtd. Itens", "Qtd. Unidades", "Valor (R$)"]
  ws_res.append([])
  ws_res.append(res_headers)
  for c_idx in range(1, 5):
    c = ws_res.cell(3, c_idx)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border

  sem_div_df = df_atual[df_atual["Ajuste"] == 0]
  faltas_df = df_atual[df_atual["Ajuste"] < 0]
  sobras_df = df_atual[df_atual["Ajuste"] > 0]
  cong_zero_df = df_atual[df_atual["Congelado"] == 0]

  val_est_col = (
      "Valor de Estoque (R$)"
      if "Valor de Estoque (R$)" in df_atual.columns
      else None
  )
  val_aju_col = (
      "Valor de Ajuste (R$)"
      if "Valor de Ajuste (R$)" in df_atual.columns
      else None
  )

  rows_resumo = [
      (
          "Total de SKUs Inventariados",
          int(len(df_atual)),
          int(df_atual["Digitado"].sum()),
          round(float(df_atual[val_est_col].sum()), 2) if val_est_col else 0.0,
      ),
      (
          "Itens sem Divergência (Ajuste = 0)",
          int(len(sem_div_df)),
          int(sem_div_df["Digitado"].sum()),
          (
              round(float(sem_div_df[val_est_col].sum()), 2)
              if val_est_col
              else 0.0
          ),
      ),
      (
          "Faltas (Ajuste Negativo)",
          int(len(faltas_df)),
          int(faltas_df["Ajuste"].sum()),
          (
              round(float(faltas_df[val_aju_col].sum()), 2)
              if val_aju_col
              else 0.0
          ),
      ),
      (
          "Sobras (Ajuste Positivo)",
          int(len(sobras_df)),
          int(sobras_df["Ajuste"].sum()),
          (
              round(float(sobras_df[val_aju_col].sum()), 2)
              if val_aju_col
              else 0.0
          ),
      ),
      (
          "Itens com Congelado = 0",
          int(len(cong_zero_df)),
          int(cong_zero_df["Digitado"].sum()),
          (
              round(float(cong_zero_df[val_est_col].sum()), 2)
              if val_est_col
              else 0.0
          ),
      ),
      (
          "Total de Itens na Aba JUSTIFICAR",
          int(len(df_just)),
          int(df_just["Ajuste"].sum()),
          (
              round(float(df_just[val_aju_col].sum()), 2)
              if val_aju_col
              else 0.0
          ),
      ),
  ]

  for idx, r_vals in enumerate(rows_resumo, start=4):
    ws_res.append(list(r_vals))
    for c_idx in range(1, 5):
      cell = ws_res.cell(idx, c_idx)
      cell.font = font_bold if idx == 9 else font_data
      cell.border = thin_border
      if idx == 9:
        cell.fill = fill_total
      elif idx % 2 == 1:
        cell.fill = fill_zebra
      if c_idx == 1:
        cell.alignment = Alignment(horizontal="left", vertical="center")
      elif c_idx in [2, 3]:
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="center", vertical="center")
      elif c_idx == 4:
        cell.number_format = "R$ #,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")

  chart = BarChart()
  chart.type = "col"
  chart.style = 10
  chart.title = "Divergências por Tipo (Qtd. de Itens)"
  chart.y_axis.title = "Qtd. Itens"
  chart.x_axis.title = "Categoria"
  data_ref = Reference(ws_res, min_col=2, min_row=5, max_row=8)
  cats_ref = Reference(ws_res, min_col=1, min_row=5, max_row=8)
  chart.add_data(data_ref, titles_from_data=False)
  chart.set_categories(cats_ref)
  chart.legend = None
  chart.width = 18
  chart.height = 12
  ws_res.add_chart(chart, "A12")

  ws_res.column_dimensions["A"].width = 36
  ws_res.column_dimensions["B"].width = 16
  ws_res.column_dimensions["C"].width = 18
  ws_res.column_dimensions["D"].width = 20

  output_buffer = io.BytesIO()
  wb.save(output_buffer)
  output_buffer.seek(0)
  return output_buffer, df_just


def renderizar_bloco_5_justificar():
  st.markdown("### 📦 AUTOMAÇÃO DE RELATÓRIO DE INVENTÁRIO (BLOCO 3)")
  st.markdown(
      "Faça o upload dos arquivos abaixo e clique no botão para gerar a"
      " planilha Excel pronta para justificar."
  )

  col1, col2 = st.columns(2)

  with col1:
    arquivo_csv = st.file_uploader(
        "1. Arquivo CSV do Inventário Atual (Obrigatório)",
        type=["csv"],
        help="Ex: 11021.23.11.2026.csv",
        key="b5_arquivo_csv",
    )

  with col2:
    arquivo_anterior = st.file_uploader(
        "2. CSV do Inventário Anterior (Opcional)",
        type=["csv", "xlsx", "xls"],
        help="Para cruzar a coluna 'Ajuste Inventario Anterior'",
        key="b5_arquivo_anterior",
    )

  st.divider()

  if st.button(
      "🚀 Gerar Relatório de Inventário",
      type="primary",
      use_container_width=True,
      key="b5_btn_gerar",
  ):
    if arquivo_csv is None:
      st.warning(
          "⚠️ Por favor, faça o upload do arquivo CSV do inventário atual"
          " primeiro."
      )
    else:
      try:
        with st.spinner("Processando arquivos e gerando planilha..."):
          df_atual = _b5_ler_csv_inventario(arquivo_csv)
          mapa_anterior = _b5_ler_inventario_anterior(arquivo_anterior)
          excel_bytes, df_just = _b5_gerar_excel_memoria(
              df_atual, mapa_anterior, arquivo_csv.name
          )

          df_preview = df_just[[
              "Código do Produto",
              "Descrição do Produto",
              "Emb",
              "Congelado",
              "Digitado",
              "Ajuste",
              "Saldo em Estoque",
          ]].copy()
          df_preview["Ajuste Inventario Anterior"] = (
              df_preview["Código do Produto"]
              .map(mapa_anterior)
              .fillna(0)
              .astype(int)
          )

        st.success(
            f"✅ Relatório gerado! **{len(df_just)} itens** filtrados para"
            f" justificativa (de {len(df_atual)} itens totais)."
        )

        nome_saida = (
            f"Relatorio_Inventario_{os.path.splitext(arquivo_csv.name)[0]}.xlsx"
        )
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=excel_bytes,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="b5_btn_download",
        )

        with st.expander(
            "👀 Pré-visualizar itens da aba JUSTIFICAR", expanded=True
        ):
          st.dataframe(df_preview, use_container_width=True, hide_index=True)

      except Exception as e:
        st.error(f"❌ Erro ao processar os arquivos: {e}")


# ====================================================================================
# APLICAÇÃO STREAMLIT PRINCIPAL
# ====================================================================================
def main():
  st.set_page_config(
      page_title="GRUPO KMB // NOVA CORE 2", page_icon="⚡", layout="wide"
  )

  SENHA_MESTRE = st.secrets.get("SENHA_MESTRE", "novacore2026")
  if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

  caminho_logo_tela = (
      "brasao_sem_fundo.png"
      if os.path.exists("brasao_sem_fundo.png")
      else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
  )

  if not st.session_state.autenticado:
    st.markdown(
        """
        <style>
        .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
    with col_l2:
      st.markdown("<br><br>", unsafe_allow_html=True)
      if caminho_logo_tela:
        st.image(caminho_logo_tela, use_container_width=True)
      else:
        st.warning("⚠️ Brasão não encontrado na pasta.")

      st.markdown("## ⚡ NOVA CORE 2 // AUTENTICAÇÃO")
      st.markdown(
          "Insira a senha de acesso executivo para desbloquear o sistema."
      )

      senha_digitada = st.text_input(
          "Senha de Acesso", type="password", key="input_senha"
      )
      if st.button("🔓 Entrar no Sistema", use_container_width=True):
        if senha_digitada == SENHA_MESTRE:
          st.session_state.autenticado = True
          st.rerun()
        else:
          st.error("❌ Senha incorreta. Acesso negado.")
    return

  st.markdown(
      """
    <style>
    .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
    .main-header { padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    </style>
    """,
      unsafe_allow_html=True,
  )

  with st.sidebar:
    if caminho_logo_tela:
      st.image(caminho_logo_tela, use_container_width=True)
    st.markdown("### **GRUPO KMB**")
    st.markdown("*Nova Core 2 • Geração de Conferência & Inventário*")
    st.markdown("---")

    modo_operacao = st.radio(
        "Modo de Operação:",
        [
            "Planilha & PDF de Conferência",
            "Gerador de Etiquetas QR Code",
            "Relatório de Inventário",
        ],
        key="radio_modo_operacao_nc2",
    )

    st.markdown("---")
    if st.button("🔒 Encerrar Sessão (Sair)", use_container_width=True):
      st.session_state.autenticado = False
      st.rerun()

  st.markdown(
      """
    <div class="main-header">
        <div class="jarvis-title">⚡ GRUPO KMB // NOVA CORE 2</div>
        <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">MÓDULO EXCLUSIVO DE PLANILHA, PDF DE CONFERÊNCIA & RELATÓRIO DE INVENTÁRIO</div>
    </div>
    """,
      unsafe_allow_html=True,
  )

  if modo_operacao == "Relatório de Inventário":
    renderizar_bloco_5_justificar()
    return

  if modo_operacao == "Gerador de Etiquetas QR Code":
    st.markdown(
        "### 🏷️ GERADOR DE ETIQUETAS QR CODE // PADRÃO A4 (14 POR PÁGINA)"
    )
    st.markdown(
        "Gere etiquetas com QR Code, numeração encorpada e brasão oficial"
        " prontas para impressão em folha A4."
    )

    col_etq1, col_etq2 = st.columns(2)
    with col_etq1:
      num_inicio = st.number_input(
          "Do número:",
          min_value=1,
          max_value=99999,
          value=1,
          step=1,
          key="etq_inicio",
      )
    with col_etq2:
      num_fim = st.number_input(
          "Ao número:",
          min_value=1,
          max_value=99999,
          value=60,
          step=1,
          key="etq_fim",
      )

    if num_inicio > num_fim:
      st.error(
          "❌ O número inicial ('Do número') não pode ser maior que o número"
          " final ('Ao número')."
      )
    else:
      total_etiquetas = (num_fim - num_inicio) + 1
      total_paginas = (total_etiquetas + 13) // 14
      st.info(
          f"📊 Serão geradas **{total_etiquetas} etiquetas** distribuídas em"
          f" **{total_paginas} página(s) A4** (2 colunas × 7 linhas)."
      )

      if st.button("⚙️ Gerar PDF de Etiquetas", use_container_width=True):
        with st.spinner("Gerando etiquetas em alta definição..."):
          pdf_etiquetas_buf = gerar_pdf_etiquetas_kmb(
              int(num_inicio), int(num_fim), caminho_logo_tela
          )
          st.session_state["pdf_etiquetas_pronto"] = (
              pdf_etiquetas_buf.getvalue()
          )
          st.session_state["pdf_etiquetas_nome"] = (
              f"Etiquetas_KMB_{int(num_inicio):02d}_a_{int(num_fim):02d}.pdf"
          )

      if "pdf_etiquetas_pronto" in st.session_state:
        st.success(
            f"✅ Arquivo **{st.session_state['pdf_etiquetas_nome']}** pronto"
            " para impressão!"
        )
        st.download_button(
            label="📥 Baixar PDF de Etiquetas",
            data=st.session_state["pdf_etiquetas_pronto"],
            file_name=st.session_state["pdf_etiquetas_nome"],
            mime="application/pdf",
            use_container_width=True,
        )
    return

  st.markdown("### 📋 GERAÇÃO AUTOMATIZADA DE LISTAS DE CONFERÊNCIA")
  st.markdown(
      "Insira o número da loja e carregue os arquivos abaixo para gerar as"
      " listas com todas as colunas preservadas em linha única:"
  )

  numero_loja_texto = st.text_input(
      "Número da Loja (ex: 13016 ou 17315)",
      value="",
      placeholder="Digite apenas números...",
      max_chars=5,
      key="numero_loja_texto_conferencia",
  )

  numero_loja_limpo = numero_loja_texto.strip()
  loja_valida = numero_loja_limpo.isdigit() and int(numero_loja_limpo) >= 1

  if not loja_valida:
    if numero_loja_limpo != "":
      st.warning(
          "⚠️ Digite um número de loja válido (apenas números, maior que zero)."
      )
    else:
      st.info(
          "ℹ️ Informe o número da loja acima para liberar o envio dos arquivos."
      )
    return

  numero_loja_ui = int(numero_loja_limpo)

  col_up1, col_up2 = st.columns(2)
  with col_up1:
    arq_loc = st.file_uploader(
        "1. Arquivo de Itens por Localização (ex: Itens_localizacao.csv)",
        type=["csv", "xlsx", "xls"],
        key="up_loc",
    )
  with col_up2:
    arq_est = st.file_uploader(
        "2. Arquivo de Saldo de Estoque (ex: saldo-estoque.csv)",
        type=["csv", "xlsx", "xls"],
        key="up_est",
    )

  if arq_loc and arq_est:
    df_loc_temp, msg1 = processar_arquivo_upload(
        arq_loc.getvalue(), arq_loc.name
    )
    df_est_temp, msg2 = processar_arquivo_upload(
        arq_est.getvalue(), arq_est.name
    )

    if df_loc_temp is not None and df_est_temp is not None:
      st.success("✅ Arquivos carregados e validados com sucesso!")

      st.markdown("---")
      st.markdown("### 👁️ Pré-visualização na Tela")

      buf_preview = gerar_planilha_filtrada_kmb(
          df_loc_temp, df_est_temp, numero_loja_ui, "TODOS"
      )
      df_preview_exibir = pd.read_excel(
          buf_preview, sheet_name=0, skiprows=3, dtype=str
      ).fillna("")

      cols_vistas = {}
      novas_colunas = []
      for c in df_preview_exibir.columns:
        c_str = str(c).strip()
        if "." in c_str and c_str.split(".")[0].isdigit():
          c_str = c_str.split(".")[0]
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
              lambda x: x[:-2] if isinstance(x, str) and x.endswith(".0") else x
          )

      st.dataframe(df_preview_exibir, use_container_width=True, height=420)

      st.download_button(
          label="📥 Baixar Pré-visualização em Excel (.xlsx)",
          data=buf_preview,
          file_name=f"Preview_SaldoLoja_Loja_{numero_loja_ui}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          use_container_width=True,
      )

      st.markdown("---")
      st.markdown("### 📥 1. Listas de Conferência em Excel (Ordem Alfabética)")
      c_imp1, c_imp2, c_imp3, c_imp4 = st.columns(4)

      with c_imp1:
        buf_zerados = gerar_planilha_filtrada_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS"
        )
        st.download_button(
            "📥 Baixar Excel (Zerados)",
            data=buf_zerados,
            file_name=f"Excel_Zerados_Loja_{numero_loja_ui}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
      with c_imp2:
        buf_faltas = gerar_planilha_filtrada_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS"
        )
        st.download_button(
            "📥 Baixar Excel (Faltas)",
            data=buf_faltas,
            file_name=f"Excel_Faltas_Loja_{numero_loja_ui}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
      with c_imp3:
        buf_sobras = gerar_planilha_filtrada_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS"
        )
        st.download_button(
            "📥 Baixar Excel (Sobras)",
            data=buf_sobras,
            file_name=f"Excel_Sobras_Loja_{numero_loja_ui}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
      with c_imp4:
        buf_sem_saldo = gerar_planilha_filtrada_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO"
        )
        st.download_button(
            "📥 Baixar Excel (Sem Saldo)",
            data=buf_sem_saldo,
            file_name=f"Excel_SemSaldo_Loja_{numero_loja_ui}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

      data_atual_str = pd.Timestamp.now().strftime("%d.%m.%Y")
      st.markdown("---")
      st.markdown("### 📄 2. Listas de Conferência em PDF (Otimizado para Papel A4)")
      c_pdf1, c_pdf2, c_pdf3, c_pdf4 = st.columns(4)

      with c_pdf1:
        pdf_zerados = gerar_pdf_conferencia_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS"
        )
        st.download_button(
            "📑 Baixar PDF (Zerados)",
            data=pdf_zerados,
            file_name=f"{numero_loja_ui}.{data_atual_str}.zerados.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
      with c_pdf2:
        pdf_faltas = gerar_pdf_conferencia_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS"
        )
        st.download_button(
            "📑 Baixar PDF (Faltas)",
            data=pdf_faltas,
            file_name=f"{numero_loja_ui}.{data_atual_str}.faltas.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
      with c_pdf3:
        pdf_sobras = gerar_pdf_conferencia_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS"
        )
        st.download_button(
            "📑 Baixar PDF (Sobras)",
            data=pdf_sobras,
            file_name=f"{numero_loja_ui}.{data_atual_str}.sobras.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
      with c_pdf4:
        pdf_sem_saldo = gerar_pdf_conferencia_kmb(
            df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO"
        )
        st.download_button(
            "📑 Baixar PDF (Sem Saldo)",
            data=pdf_sem_saldo,
            file_name=f"{numero_loja_ui}.{data_atual_str}.sem_saldo.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    else:
      st.error("❌ Erro ao processar os arquivos.")


if __name__ == "__main__":
  main()
