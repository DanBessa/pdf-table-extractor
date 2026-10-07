from datetime import datetime
import os
import re
from tkinter import filedialog
import pandas as pd
import pdfplumber
import traceback

ANO_FALLBACK = datetime.now().year


def limpar_e_converter_valor(val_str: str):
  """Converte o valor string do Santander para float e define se é C ou D."""
  if not val_str:
    return 0.0, "C"

  v = str(val_str).strip()
  eh_debito = "-" in v or v.endswith("D") or v.startswith("-")

  v_limpo = re.sub(r"[^\d,\.]", "", v)
  if "," in v_limpo and "." in v_limpo:
    v_limpo = v_limpo.replace(".", "").replace(",", ".")
  elif "," in v_limpo:
    v_limpo = v_limpo.replace(",", ".")

  try:
    val_float = float(v_limpo)
    if eh_debito and val_float > 0:
      val_float = -val_float

    tipo = "D" if val_float < 0 else "C"
    return val_float, tipo
  except ValueError:
    return 0.0, "C"


def extrair_ano_extrato(pdf) -> int:
  meses = (
      "janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro"
  )
  for page in pdf.pages[:3]:
    texto = page.extract_text() or ""
    m = re.search(rf"(?:{meses})/(\d{{4}})", texto, re.IGNORECASE)
    if m:
      return int(m.group(1))
  return ANO_FALLBACK


def extrair_santander(caminho_pdf: str) -> pd.DataFrame:
  transacoes = []
  ignorar_termos = [
      "EXTRATO CONSOLIDADO",
      "SANTANDER EMPRESAS",
      "SALDO EM",
      "SALDO ANTERIOR",
      "SALDO FINAL",
      "SALDO TOTAL",
      "SALDO DISPONIVEL",
      "DÉBITOS CRÉDITOS",
      "DEBITOS CREDITOS",
      "DATA DESCRIÇÃO",
      "DATA DESCRICAO",
      "CONTINUAÇÃO",
      "CONTINUACAO",
      "PAGINA:",
      "PÁGINA:",
      "RESUMO -",
      "CENTRAL DE ATENDIMENTO",
      "BANCO SANTANDER",
      "SAC",
      "OUVIDORIA",
      "LIMITES CONTRATADOS",
      "SALDOS POR PERÍODO",
      "SALDOS POR PERIODO",
  ]
  
  # Palavras-chave que indicam o início de um novo bloco de transação
  start_markers = (
      "PIX RECEBID", "PIX ENVIAD", "TED RECEB", "TED ENVIAD", 
      "PAGAMENTO DE BOLETO", "PAGAMENTO CARTAO", "PAGAMENTO DE TITULO", "PAGAMENTO DARF",
      "PAGTO ELETRONICO", "TARIFA", "DEBITO AUTOM", "DÉBITO AUTOM",
      "APLICACAO", "APLICAÇÃO", "RESGATE", "DOC RECEBID", "DOC ENVIAD",
      "TRANSFERENCIA", "TRANSF ", "SAQUE", "DEPOSITO", "ENCARGOS",
      "JUROS", "IOF", "MANUTENCAO", "MANUTENÇÃO", "MENSALIDADE"
  )

  try:
    with pdfplumber.open(caminho_pdf) as pdf:
      ano_extrato = extrair_ano_extrato(pdf)
      data_atual = None
      em_conta_corrente = False
      
      # Máquina de estado para armazenar a transação atual enquanto coleta as linhas
      current_tx = {"linhas": [], "valor": 0.0, "tipo": "", "data": ""}

      def save_current_tx():
          if current_tx["valor"] == 0.0 and not current_tx["linhas"]:
              return
          
          # Une todas as linhas capturadas para a transação
          hist_raw = " - ".join(current_tx["linhas"]).strip()
          
          # ESTRATÉGIA MESTRA: Deleta termos bancários genéricos para sobrar SOMENTE o fornecedor/Getnet
          termos_remover = [
              r"PAGAMENTO DE BOLETO OUTROS BANCOS",
              r"PAGAMENTO DE BOLETO",
              r"PAGAMENTO CARTAO DE DEBITO",
              r"PAGAMENTO CARTAO DE CREDITO",
          ]
          hist_completo = hist_raw
          for t in termos_remover:
              hist_completo = re.sub(rf"{t}", "", hist_completo, flags=re.IGNORECASE).strip()

          # Limpa os traços e lixos que sobraram após apagar o termo genérico
          hist_completo = re.sub(r"-\s*-", "-", hist_completo)
          hist_completo = re.sub(r"^[-\s]+", "", hist_completo)
          hist_completo = re.sub(r"[-\s]+$", "", hist_completo)

          if not hist_completo:
              hist_completo = hist_raw if hist_raw else "LANCAMENTO BANCARIO"

          # Validação Semântica para Segurança
          h_up = hist_raw.upper()
          val_num = current_tx["valor"]
          tipo = current_tx["tipo"]
          
          if any(k in h_up for k in ["RECEBIDO", "RECEBIDA", "RESGATE", "ESTORNO", "GETNET"]):
            val_num = abs(val_num)
            tipo = "C"
          elif any(k in h_up for k in ["TARIFA", "PAGAMENTO DE BOLETO", "PAGAMENTO DARF", "PGTO TRIBUTO", "ENVIADO", "ENVIADA", "APLICACAO", "APLICAÇÃO", "COMPRA", "DEBITO", "DÉBITO", "DEVOLVIDO"]):
            if "GETNET" not in h_up:
              val_num = -abs(val_num)
              tipo = "D"

          if val_num != 0.0:
              transacoes.append({
                  "Data": current_tx["data"],
                  "Histórico": hist_completo,
                  "Valor": val_num,
                  "Tipo": tipo,
              })

      for page in pdf.pages:
        words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
        if not words:
          continue

        words_ordenadas = sorted(words, key=lambda w: (w["top"], w["x0"]))
        linhas_visuais = []
        for w in words_ordenadas:
          if not linhas_visuais or abs(w["top"] - linhas_visuais[-1]["top"]) > 3.5:
            linhas_visuais.append({"top": w["top"], "words": [w]})
          else:
            linhas_visuais[-1]["words"].append(w)

        for item_linha in linhas_visuais:
          palavras = item_linha["words"]
          linha_completa = " ".join(w["text"] for w in palavras).strip()
          linha_upper = linha_completa.upper()

          if "CONTA CORRENTE" in linha_upper or "MOVIMENTA" in linha_upper:
            em_conta_corrente = True
            continue
          if any(sec in linha_upper for sec in ["SALDOS POR PERÍODO", "SALDOS POR PERIODO", "DÉBITO AUTOMÁTICO", "COMPRAS COM CARTÃO DE DÉBITO", "CRÉDITOS CONTRATADOS", "INVESTIMENTOS", "CONTAMAX EMPRESARIAL"]):
            em_conta_corrente = False
            continue
          if not em_conta_corrente:
            continue
          if any(termo in linha_upper for termo in ignorar_termos):
            continue

          # Separa colunas estritamente
          txt_left = " ".join(w["text"] for w in palavras if w["x0"] < 365).strip()
          txt_credito = " ".join(w["text"] for w in palavras if 365 <= w["x0"] < 425).strip()
          txt_debito = " ".join(w["text"] for w in palavras if 425 <= w["x0"] < 500).strip()

          m_data = re.search(r"^(\d{2}/\d{2}(?:/\d{2,4})?)\b", txt_left)
          tem_data = bool(m_data)
          
          if tem_data:
            d_raw = m_data.group(1)
            dia, mes = d_raw.split("/")[:2]
            data_atual = f"{dia.zfill(2)}/{mes.zfill(2)}/{ano_extrato}"
            txt_desc_linha = txt_left[len(d_raw):].strip() # Tira a data da descrição
          else:
            txt_desc_linha = txt_left

          txt_desc_linha = re.sub(r"^-+|-+$", "", txt_desc_linha).strip()

          m_cred = re.search(r"(-?\s*[\d\.]*,\d{2}\s*[-CD]?|\b[\d\.]*,\d{2}\s*-$)", txt_credito)
          m_deb = re.search(r"(-?\s*[\d\.]*,\d{2}\s*[-CD]?|\b[\d\.]*,\d{2}\s*-$)", txt_debito)
          tem_valor = bool(m_cred or m_deb)

          # Verifica se a linha indica o começo de um NOVO bloco de transação
          is_new_tx = False
          if tem_data:
            is_new_tx = True
          elif any(txt_desc_linha.upper().startswith(mk) for mk in start_markers):
            is_new_tx = True

          # Se é um novo bloco, salva a transação anterior completa e reseta o pacote
          if is_new_tx and len(current_tx["linhas"]) > 0:
            save_current_tx()
            current_tx = {"linhas": [], "valor": 0.0, "tipo": "", "data": data_atual}

          if tem_data:
             current_tx["data"] = data_atual
          elif not current_tx["data"]:
             current_tx["data"] = data_atual

          if txt_desc_linha:
            current_tx["linhas"].append(txt_desc_linha)

          # Anexa o valor e tipo se houver na linha
          if tem_valor:
            val_str = m_deb.group(0).strip() if m_deb else m_cred.group(0).strip()
            v, t = limpar_e_converter_valor(val_str)
            if m_deb and not m_cred:
              t = "D"
              v = -abs(v)
            else:
              t = "C"
              v = abs(v)
            current_tx["valor"] = v
            current_tx["tipo"] = t

      # Ao terminar de ler as páginas, salva o último bloco que ficou pendente
      if len(current_tx["linhas"]) > 0:
          save_current_tx()

    if not transacoes:
      return None

    df = pd.DataFrame(transacoes)[["Data", "Histórico", "Valor", "Tipo"]]
    return df

  except Exception as e:
    print(f"Erro ao processar PDF Santander: {e}")
    traceback.print_exc()
    raise e


def iniciar_processamento(pdf_path=None):
  if not pdf_path:
    pdf_path = filedialog.askopenfilename(
        title="Selecione o extrato (Santander - Modelo 1)",
        filetypes=[("Arquivos PDF", "*.pdf")],
    )
  if not pdf_path:
    raise UserWarning("Nenhum arquivo selecionado.")

  try:
    df_transacoes = extrair_santander(pdf_path)

    if df_transacoes is None or df_transacoes.empty:
      raise UserWarning(
          f"Nenhuma transação válida foi encontrada em"
          f" {os.path.basename(pdf_path)}."
      )

    total_cred = len(df_transacoes[df_transacoes["Tipo"] == "C"])
    total_deb = len(df_transacoes[df_transacoes["Tipo"] == "D"])
    print(f"✅ Santander Modelo 1: {len(df_transacoes)} lançamentos extraídos.")
    print(f"   🟢 Entradas (+ / C): {total_cred}")
    print(f"   🔴 Saídas (- / D): {total_deb}")

    caminho_saida = os.path.splitext(pdf_path)[0] + ".xlsx"
    df_transacoes.to_excel(caminho_saida, index=False)
    return caminho_saida

  except Exception as e:
    traceback.print_exc()
    raise Exception(f"Erro ao processar '{os.path.basename(pdf_path)}':\n{e}")


if __name__ == "__main__":
  iniciar_processamento()