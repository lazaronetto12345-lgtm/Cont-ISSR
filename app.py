# ============================================================
#  APP SSR — v28.5.0 (Géis com Marcações + Auto-Pastas + Importação com Nomes)
#  ✅ LOGIN COM USUÁRIO E SENHA PADRÃO: ifesbiomol / biomol102030
#  ✅ EXCEL EM BLOCOS MULTIPRIMER: Arial 11, Centralizado, Sem Cores de Fundo
#  ✅ MODO CRIAÇÃO LIVRE DE COLUNAS (clicar, arrastar, deletar)
#  ✅ Zoom com roda, Pan com meio/Espaço, UPGMA Multilocus e Excel
#  ✅ Salva Géis de Conferência em JPG + Backup Físico em Qualquer PC
# ============================================================

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import linkage, dendrogram
from scipy.spatial.distance import squareform
from scipy.ndimage import uniform_filter1d
import cv2
import io
import base64
import os
import tempfile
import hashlib
import pickle
import openpyxl
from datetime import datetime

# ============================================================
#  🔑 CREDENCIAIS PADRÃO DE ACESSO
# ============================================================
USUARIO_PADRAO = "ifesbiomol"
SENHA_PADRAO = "biomol102030"


st.set_page_config(page_title="SSR Pro v28.5", page_icon="🧬", layout="wide")
st.markdown("""
<style>
    .block-container { padding-top:1rem; padding-bottom:1rem; max-width:100% !important; }
    iframe { border:none !important; }
    div[data-testid="column"] button { width:100%; margin-bottom:4px; }
</style>
""", unsafe_allow_html=True)


# ============================================================
#  🔒 SISTEMA DE VERIFICAÇÃO DE USUÁRIO E SENHA
# ============================================================

def verificar_autenticacao():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if st.session_state["autenticado"]:
        return True

    # Tela de Login / Autenticação
    st.markdown("<br><br>", unsafe_allow_html=True)
    c_left, c_center, c_right = st.columns([1, 2, 1])
    with c_center:
        st.title("🔒 Acesso Restrito — SSR Pro")
        st.info("Entre com suas credenciais de acesso para acessar o sistema.")
        
        usuario_digitado = st.text_input("Usuário:", key="input_usuario_acesso")
        senha_digitada = st.text_input("Senha:", type="password", key="input_senha_acesso")
        
        if st.button("🔑 Entrar no Sistema", type="primary", use_container_width=True):
            if usuario_digitado == USUARIO_PADRAO and senha_digitada == SENHA_PADRAO:
                st.session_state["autenticado"] = True
                st.success("✅ Acesso liberado!")
                st.rerun()
            else:
                st.error("❌ Usuário ou senha incorretos!")
                
    return False

# Bloqueia a execução caso as credenciais não estejam corretas
if not verificar_autenticacao():
    st.stop()


# ============================================================
#  📁 SISTEMA DE BACKUP LOCAL, SALVAMENTO DE GÉIS E IMPORTAÇÃO
# ============================================================
BACKUP_DIR = "ssr_resultados_salvos"
BACKUP_FILE = os.path.join(BACKUP_DIR, "backup_session_state.pkl")
GEIS_DIR = os.path.join(BACKUP_DIR, "geis")

def garantir_pastas():
    """Garante que as pastas de backup e de géis existam em qualquer computador"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    os.makedirs(GEIS_DIR, exist_ok=True)
    return os.path.abspath(BACKUP_DIR), os.path.abspath(GEIS_DIR)

def _nome_arquivo_seguro(texto):
    t = str(texto).strip()
    return "".join(c if (c.isalnum() or c in " _-") else "_" for c in t).strip().replace(" ", "_") or "primer"

def salvar_gel_conferencia(img_bgr, nome_primer, bandas_salvas=None, lista_f1=None, lista_f2=None,
                            edges1=None, edges2=None, w1=0, skip1=0, skip2=0, y_guia=0):
    """
    Gera e salva no computador a imagem do gel processado com:
    - Linha do Laser (Verde)
    - Linhas horizontais das bandas marcadas
    - Retângulos verdes das presenças (1)
    """
    try:
        _, geis_abs = garantir_pastas()
        if img_bgr is None:
            return False, "Imagem base do gel não encontrada."

        img = img_bgr.copy()
        h, w = img.shape[:2]
        lista_f1 = lista_f1 or []
        lista_f2 = lista_f2 or []
        bandas_salvas = bandas_salvas or []
        edges1 = edges1 or []
        edges2 = edges2 or []

        # 1. Linha do Laser
        if y_guia and 0 < int(y_guia) < h:
            cv2.line(img, (0, int(y_guia)), (w - 1, int(y_guia)), (0, 255, 120), 2)

        # 2. Desenhar linhas das bandas e letras identificadoras
        for i, b in enumerate(bandas_salvas):
            y = int(round(float(b.get("y", -1))))
            if 0 <= y < h:
                cv2.line(img, (0, y), (w - 1, y), (200, 200, 200), 1)
                letra = chr(97 + i) if i < 26 else f"b{i}"
                cv2.putText(img, letra, (10, max(20, y - 4)), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 255), 2)

        # 3. Desenhar retângulos verdes das presenças marcadas
        def _centro_lane(edges, x_off, lane):
            if not edges or lane < 0 or lane >= len(edges) - 1:
                return None, None
            x0 = float(edges[lane]) + x_off
            x1 = float(edges[lane + 1]) + x_off
            return (x0 + x1) / 2.0, max(6.0, (x1 - x0) * 0.75)

        for b in bandas_salvas:
            y = int(round(float(b.get("y", -1))))
            marks = b.get("marks") or []
            if y < 0 or y >= h:
                continue
            for col_idx, m in enumerate(marks):
                if int(m) != 1:
                    continue
                cx, tw = None, 20
                if col_idx < len(lista_f1):
                    lane = col_idx + int(skip1)
                    cx, tw = _centro_lane(edges1, 0, lane)
                else:
                    idx2 = col_idx - len(lista_f1)
                    if 0 <= idx2 < len(lista_f2):
                        lane = idx2 + int(skip2)
                        cx, tw = _centro_lane(edges2, float(w1), lane)
                if cx is None:
                    continue
                x1 = int(cx - tw / 2)
                x2 = int(cx + tw / 2)
                y1 = max(0, y - 3)
                y2 = min(h - 1, y + 3)
                cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), -1)
                cv2.rectangle(img, (x1, y1), (x2, y2), (0, 80, 0), 1)

        safe = _nome_arquivo_seguro(nome_primer)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho = os.path.join(GEIS_DIR, f"{safe}_{ts}.jpg")
        ok = cv2.imwrite(caminho, img, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
        if not ok:
            return False, "Erro ao gravar o arquivo JPG."
        return True, os.path.abspath(caminho)
    except Exception as e:
        return False, str(e)

def salvar_backup_local():
    """Grava as matrizes salvas em um arquivo físico no computador e gera o Excel atualizado"""
    try:
        pasta_abs, _ = garantir_pastas()
        
        # 1. Salva o dicionário de primers para recuperação automática do app
        with open(BACKUP_FILE, "wb") as f:
            pickle.dump(st.session_state["todas_matrizes"], f)
        
        # 2. Gera também a planilha Excel em tempo real direto na pasta do computador
        if st.session_state["todas_matrizes"]:
            excel_data = exportar_excel_completo(
                st.session_state["todas_matrizes"], 
                nome_export="Backup_Automatico"
            )
            excel_path = os.path.join(BACKUP_DIR, "SSR_Matriz_Backup_Completo.xlsx")
            with open(excel_path, "wb") as f_excel:
                f_excel.write(excel_data)
                
        return True, pasta_abs
    except Exception as e:
        return False, str(e)

def carregar_backup_local():
    """Busca se há algum backup no disco e recupera para a memória ao iniciar"""
    if os.path.exists(BACKUP_FILE):
        try:
            with open(BACKUP_FILE, "rb") as f:
                dados = pickle.load(f)
                if isinstance(dados, dict):
                    st.session_state["todas_matrizes"] = dados
                    return True, len(dados)
        except Exception:
            pass
    return False, 0

def build_df_from_bands(bands, acessos, primer_name):
    row_names = []
    data_rows = []
    for b in bands:
        b_label = b["band"]
        if b_label.startswith(f"{primer_name}_"):
            idx_name = b_label
        else:
            idx_name = f"{primer_name}_{b_label}"
        row_names.append(idx_name)
        data_rows.append(b["vals"])
        
    df = pd.DataFrame(data_rows, index=row_names, columns=acessos)
    return df

def importar_excel_completo(file_bytes):
    """Lê um arquivo Excel gerado pelo próprio App e o reconstrói em DataFrames válidos"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_name = None
        for name in wb.sheetnames:
            if name.startswith("Matriz_"):
                sheet_name = name
                break
        if not sheet_name:
            sheet_name = wb.sheetnames[0]
            
        ws = wb[sheet_name]
        
        first_row = [cell.value for cell in ws[1]]
        if not first_row or len(first_row) < 2:
            return None, "A planilha parece estar vazia ou fora do formato padrão."
            
        first_primer = str(first_row[0]).strip()
        acessos = [str(val).strip() for val in first_row[1:] if val is not None]
        
        primers_dict = {}
        current_primer = first_primer
        current_bands = []
        
        for r_idx in range(2, ws.max_row + 1):
            row_cells = [ws.cell(row=r_idx, column=c_idx) for c_idx in range(1, len(acessos) + 2)]
            col1_val = row_cells[0].value
            vals = [c.value for c in row_cells[1:]]
            
            if col1_val is None and all(v is None for v in vals):
                continue
                
            if col1_val is not None and all(v is None for v in vals):
                if current_bands:
                    df = build_df_from_bands(current_bands, acessos, current_primer)
                    primers_dict[current_primer] = df
                current_primer = str(col1_val).strip()
                current_bands = []
                continue
                
            if col1_val is not None:
                band_name = str(col1_val).strip()
                binary_vals = []
                for v in vals:
                    if v is None:
                        binary_vals.append(False)
                    else:
                        try:
                            binary_vals.append(bool(int(float(v))))
                        except:
                            binary_vals.append(False)
                
                if len(binary_vals) < len(acessos):
                    binary_vals += [False] * (len(acessos) - len(binary_vals))
                else:
                    binary_vals = binary_vals[:len(acessos)]
                    
                current_bands.append({"band": band_name, "vals": binary_vals})
                
        if current_bands:
            df = build_df_from_bands(current_bands, acessos, current_primer)
            primers_dict[current_primer] = df
            
        return primers_dict, None
    except Exception as e:
        return None, str(e)


# ============================================================
#  FUNÇÕES PYTHON DE ORDENAÇÃO E PROCESSAMENTO
# ============================================================

def ordenar_ids(lista):
    """
    Ordena numericamente os IDs selecionados (ex: 1, 7, 8, 15, 22, 100...).
    Marcadores 'L' ou 'M' ficam no início.
    """
    if not lista:
        return []
    
    def chave_ordenacao(item):
        val = str(item).strip().upper()
        if val in ["L", "M", "LADDER", "MARCADOR"]:
            return (-1, 0, val)
        try:
            return (0, int(val), "")
        except ValueError:
            return (1, 0, val)
            
    return sorted(list(dict.fromkeys(lista)), key=chave_ordenacao)

def sync_e_ordenar_f1():
    """Callback para ordenar visualmente a Foto 1 se a opção estiver ativa"""
    if "f1_select" in st.session_state:
        if st.session_state.get("ordenar_ids_auto", True):
            st.session_state["f1_select"] = ordenar_ids(st.session_state["f1_select"])

def sync_e_ordenar_f2():
    """Callback para ordenar visualmente a Foto 2 se a opção estiver ativa"""
    if "f2_select" in st.session_state:
        if st.session_state.get("ordenar_ids_auto", True):
            st.session_state["f2_select"] = ordenar_ids(st.session_state["f2_select"])

def remover_sujeira(img_bgr, nivel="Desligado"):
    if nivel == "Desligado" or img_bgr is None: return img_bgr
    img = img_bgr.copy()
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    if nivel == "Leve":   thr_dark,thr_bright,area_max,max_dim,inpaint_r = 28,32,18,7,2
    elif nivel == "Forte": thr_dark,thr_bright,area_max,max_dim,inpaint_r = 16,20,70,14,3
    else:                  thr_dark,thr_bright,area_max,max_dim,inpaint_r = 22,26,40,10,2
    fundo = cv2.GaussianBlur(gray,(31,31),0)
    dark = cv2.subtract(fundo,gray); _,mask_dark = cv2.threshold(dark,thr_dark,255,cv2.THRESH_BINARY)
    bright = cv2.subtract(gray,fundo); _,mask_bright = cv2.threshold(bright,thr_bright,255,cv2.THRESH_BINARY)
    mask = cv2.bitwise_or(mask_dark,mask_bright)
    mask = cv2.morphologyEx(mask,cv2.MORPH_OPEN,np.ones((2,2),np.uint8),iterations=1)
    num,labels,stats,_ = cv2.connectedComponentsWithStats(mask,connectivity=8)
    mask_final = np.zeros_like(mask)
    for i in range(1,num):
        area=stats[i,cv2.CC_STAT_AREA]; w=stats[i,cv2.CC_STAT_WIDTH]; h=stats[i,cv2.CC_STAT_HEIGHT]
        aspect=(max(w,h)/max(1,min(w,h)))
        if area<=area_max and max(w,h)<=max_dim and aspect<=3.5: mask_final[labels==i]=255
    if np.count_nonzero(mask_final)==0: return img
    mask_final=cv2.dilate(mask_final,np.ones((2,2),np.uint8),iterations=1)
    return cv2.inpaint(img,mask_final,inpaintRadius=inpaint_r,flags=cv2.INPAINT_TELEA)

def aplicar_filtro_bw(img_bgr, modo_filtro, brilho, contraste):
    gray = cv2.cvtColor(img_bgr,cv2.COLOR_BGR2GRAY)
    if modo_filtro == "Preto e Branco (Invertido - Fundo Branco)":
        proc = cv2.cvtColor(cv2.bitwise_not(gray),cv2.COLOR_GRAY2BGR)
    elif modo_filtro == "Preto e Branco (Fundo Preto)":
        proc = cv2.cvtColor(gray,cv2.COLOR_GRAY2BGR)
    else: proc = img_bgr.copy()
    return cv2.convertScaleAbs(proc,alpha=contraste,beta=int((brilho-1.0)*50))

def transformar_imagem(img, angulo, escala, offset_y, cor_fundo):
    h,w = img.shape[:2]
    M = cv2.getRotationMatrix2D((w//2,h//2),angulo,escala)
    M[1,2] += offset_y
    return cv2.warpAffine(img,M,(w,h),borderValue=cor_fundo)

def detectar_fundo_pocos(img_bgr):
    gray=cv2.cvtColor(img_bgr,cv2.COLOR_BGR2GRAY); h,w=gray.shape
    y_max_busca=max(30,int(h*0.28)); roi=gray[:y_max_busca,:]
    roi_eq=cv2.equalizeHist(roi); roi_blur=cv2.GaussianBlur(roi_eq,(5,5),0)
    sobel=cv2.Sobel(roi_blur,cv2.CV_64F,0,1,ksize=3); sobel=np.abs(sobel)
    perfil=sobel.sum(axis=1); perfil=uniform_filter1d(perfil.astype(float),size=7)
    y0=max(5,int(h*0.03)); y1=y_max_busca-2
    if y1<=y0+5: return int(h*0.10)
    trecho=perfil[y0:y1]; limiar=np.percentile(trecho,75); candidatos=[]
    for i in range(2,len(trecho)-2):
        if trecho[i]>=limiar and trecho[i]>=trecho[i-1] and trecho[i]>=trecho[i+1]:
            candidatos.append((trecho[i],y0+i))
    if not candidatos: return int(y0+np.argmax(trecho))
    candidatos.sort(reverse=True,key=lambda t:t[0]); top=candidatos[:min(5,len(candidatos))]
    y_fundo=max(top,key=lambda t:t[1])[1]; return int(min(h-1,y_fundo+2))

def auto_calibrar(img1_bgr, img2_bgr, rot1=0.0, rot2=0.0, escala2=1.0):
    cor=(0,0,0)
    i1=transformar_imagem(img1_bgr,rot1,1.0,0,cor); i2=transformar_imagem(img2_bgr,rot2,escala2,0,cor)
    h1,w1=i1.shape[:2]; h2,w2=i2.shape[:2]
    if h2!=h1: nw2=int(w2*(h1/h2)); i2=cv2.resize(i2,(nw2,h1),interpolation=cv2.INTER_AREA)
    y1=detectar_fundo_pocos(i1); y2=detectar_fundo_pocos(i2)
    return int(y1-y2),int(y1),y1,y2

def calcular_jaccard(matriz):
    n=matriz.shape[0]; dist=np.zeros((n,n))
    for i in range(n):
        for j in range(i+1,n):
            a=int(np.sum((matriz[i]==1)&(matriz[j]==1)))
            b=int(np.sum((matriz[i]==1)&(matriz[j]==0)))
            c=int(np.sum((matriz[i]==0)&(matriz[j]==1)))
            d=a+b+c; dist[i,j]=0.0 if d==0 else 1.0-(a/d); dist[j,i]=dist[i,j]
    return dist

def fazer_upgma(dist_matrix):
    dm=dist_matrix.copy(); np.fill_diagonal(dm,0); dm=(dm+dm.T)/2
    return linkage(squareform(dm,checks=False),method="average")

def plotar_dendrograma(Z, labels, titulo):
    fig,ax=plt.subplots(figsize=(max(10,len(labels)*0.35),7))
    dendrogram(Z,labels=labels,ax=ax,leaf_rotation=90,leaf_font_size=8,
               color_threshold=0.7*max(Z[:,2]) if len(Z)>0 else 0)
    ax.set_title(titulo,fontsize=13,fontweight="bold"); ax.set_ylabel("Dissimilaridade de Jaccard")
    plt.tight_layout(); return fig

def carregar_imagem(uploaded):
    try:
        data=uploaded.read(); img=cv2.imdecode(np.frombuffer(data,np.uint8),cv2.IMREAD_COLOR)
        return img,None
    except Exception as e: return None,str(e)

# ============================================================
#  EXPORTAÇÃO EXCEL EM BLOCOS (ARIAL 11, CENTRALIZADO, SEM CORES)
# ============================================================

def exportar_excel_completo(primers_dict, nome_export="Combinado", dist_jaccard=None, acessos=None):
    from openpyxl.styles import Font, Alignment, PatternFill
    
    buf = io.BytesIO()

    if isinstance(primers_dict, pd.DataFrame):
        primers_list = [(nome_export, primers_dict)]
    elif isinstance(primers_dict, dict):
        primers_list = [(p_name, df) for p_name, df in primers_dict.items()]
    elif isinstance(primers_dict, list):
        primers_list = primers_dict
    else:
        primers_list = [(nome_export, primers_dict)]

    all_acessos = []
    for _, df_p in primers_list:
        for col in df_p.columns:
            if col not in all_acessos:
                all_acessos.append(col)

    if st.session_state.get("ordenar_ids_auto", True):
        all_acessos = ordenar_ids(all_acessos)

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        sheet_matriz = f"Matriz_{nome_export}"[:31]
        
        df_init = pd.DataFrame(columns=all_acessos)
        df_init.to_excel(w, sheet_name=sheet_matriz, index=False)
        ws1 = w.sheets[sheet_matriz]
        
        ws1.delete_rows(1, ws1.max_row + 1)

        font_arial = Font(name="Arial", size=11, bold=False)
        font_arial_bold = Font(name="Arial", size=11, bold=True)
        alinhamento_centro = Alignment(horizontal="center", vertical="center")

        current_row = 1

        for idx_primer, (p_name, df_p) in enumerate(primers_list):
            if idx_primer == 0:
                cell_a1 = ws1.cell(row=current_row, column=1, value=p_name)
                cell_a1.font = font_arial_bold
                cell_a1.alignment = alinhamento_centro

                for c_idx, acc in enumerate(all_acessos, start=2):
                    cell = ws1.cell(row=current_row, column=c_idx, value=acc)
                    cell.font = font_arial_bold
                    cell.alignment = alinhamento_centro
                
                current_row += 1
            else:
                cell_p = ws1.cell(row=current_row, column=1, value=p_name)
                cell_p.font = font_arial_bold
                cell_p.alignment = alinhamento_centro
                current_row += 1

            for band_name in df_p.index:
                s_band = str(band_name)
                band_label = s_band.split("_")[-1] if "_" in s_band else s_band
                
                cell_b = ws1.cell(row=current_row, column=1, value=band_label)
                cell_b.font = font_arial_bold
                cell_b.alignment = alinhamento_centro

                for c_idx, acc in enumerate(all_acessos, start=2):
                    val = 0
                    if acc in df_p.columns:
                        v = df_p.loc[band_name, acc]
                        val = 1 if bool(v) else 0
                    
                    cell_v = ws1.cell(row=current_row, column=c_idx, value=val)
                    cell_v.font = font_arial
                    cell_v.alignment = alinhamento_centro
                
                current_row += 1

            current_row += 1

        for col in ws1.columns:
            ml = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            ws1.column_dimensions[col_letter].width = max(ml + 3, 6)

        if dist_jaccard is not None and acessos is not None:
            df_dist = pd.DataFrame(dist_jaccard, index=acessos, columns=acessos)
            df_dist.to_excel(w, sheet_name="Dissimilaridade_Jaccard")
            ws2 = w.sheets["Dissimilaridade_Jaccard"]
            ws2.cell(row=1, column=1, value="Matriz")

            for row in ws2.iter_rows():
                for cell in row:
                    cell.font = font_arial
                    cell.alignment = alinhamento_centro
                    cell.fill = PatternFill(fill_type=None)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.000'

            for cell in ws2[1]:
                cell.font = font_arial_bold
            for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, min_col=1, max_col=1):
                row[0].font = font_arial_bold

            for col in ws2.columns:
                ml = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws2.column_dimensions[col_letter].width = max(ml + 3, 8)

    return buf.getvalue()


# ============================================================
#  HTML CANVAS v28.1
# ============================================================

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<script>
(function(){
    var S={
        setComponentReady:function(){ window.parent.postMessage({isStreamlitMessage:true,type:"streamlit:componentReady",apiVersion:1},"*"); },
        setFrameHeight:function(h){ window.parent.postMessage({isStreamlitMessage:true,type:"streamlit:setFrameHeight",height:h},"*"); },
        setComponentValue:function(v){ window.parent.postMessage({isStreamlitMessage:true,type:"streamlit:setComponentValue",value:v},"*"); },
        events:{addEventListener:function(type,cb){ window.addEventListener("message",function(ev){ if(ev.data&&ev.data.type===type) cb({detail:ev.data}); }); }}
    };
    window.Streamlit=S;
})();
</script>
<style>
*{box-sizing:border-box;}
body{margin:0;padding:0;background:#111;font-family:'Segoe UI',sans-serif;user-select:none;overflow:hidden;}
#calib-panel{background:linear-gradient(135deg,#1a252f,#2c3e50);border-bottom:2px solid #3498db;padding:10px 16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;font-size:12px;color:#ecf0f1;}
.btn-edit{padding:7px 16px;border:2px solid #3498db;border-radius:6px;background:#3498db;color:#fff;font-weight:bold;font-size:12px;cursor:pointer;white-space:nowrap;}
.btn-edit:hover{background:#2980b9;}
.btn-edit.ativo{background:#e74c3c;border-color:#e74c3c;animation:pulse 1.5s infinite;}
.btn-clear{padding:7px 16px;border:2px solid #e74c3c;border-radius:6px;background:#c0392b;color:#fff;font-weight:bold;font-size:12px;cursor:pointer;white-space:nowrap;}
.btn-clear:hover{background:#e74c3c;}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.8}}
#calib-instrucao{flex:1;padding:6px 12px;background:rgba(52,152,219,0.15);border:1px solid rgba(52,152,219,0.4);border-radius:5px;color:#3498db;font-weight:600;font-size:12px;min-height:30px;display:flex;align-items:center;}
#toolbar{padding:7px 14px;background:linear-gradient(135deg,#2c3e50,#34495e);color:#ecf0f1;border-bottom:2px solid #1a252f;font-size:11px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:5px;}
.tb-group{display:flex;align-items:center;gap:7px;flex-wrap:wrap;}
.ci{display:flex;align-items:center;gap:4px;padding:3px 8px;background:rgba(255,255,255,0.08);border-radius:4px;}
.lbl{font-weight:600;color:#3498db;font-size:11px;}
.val{color:#ecf0f1;font-size:11px;}
#container-wrapper{width:100%;height:520px;overflow:auto;background:#0d0d0d;}
#gel-canvas{cursor:crosshair;display:block;}
#container-wrapper::-webkit-scrollbar{width:10px;height:10px;}
#container-wrapper::-webkit-scrollbar-track{background:#2c3e50;}
#container-wrapper::-webkit-scrollbar-thumb{background:#3498db;border-radius:5px;}
#toast{position:fixed;bottom:14px;left:50%;transform:translateX(-50%);color:#fff;padding:8px 18px;border-radius:6px;font-size:12px;box-shadow:0 3px 10px rgba(0,0,0,0.4);z-index:9999;display:none;font-weight:600;pointer-events:none;}
</style>
</head>
<body>

<div id="calib-panel">
    <button id="btn-edit" class="btn-edit" onclick="toggleEditLanes()">📐 Ajustar / Criar Colunas Manualmente</button>
    <button id="btn-clear" class="btn-clear" onclick="clearLanes()">🗑️ Apagar Todas as Linhas</button>
    <div id="calib-instrucao">
        <strong>Zoom:</strong> roda do mouse · <strong>Mover:</strong> botão do meio OU Espaço+arrastar · <strong>Colunas:</strong> ative o modo azul
    </div>
</div>

<div id="toolbar">
    <div class="tb-group">
        <div class="ci"><span>🔄</span><span class="lbl">Roda:</span><span class="val">Zoom</span></div>
        <div class="ci"><span>🖐️</span><span class="lbl">Meio/Espaço:</span><span class="val">Mover</span></div>
        <div class="ci"><span>🔵</span><span class="lbl">Esq.(vazio):</span><span class="val">Criar banda</span></div>
        <div class="ci"><span style="color:#0f0;font-weight:bold;">━</span><span class="lbl">Esq.(banda):</span><span class="val">Marcar</span></div>
        <div class="ci"><span>🔴</span><span class="lbl">Dir.:</span><span class="val">Excluir banda</span></div>
    </div>
    <div style="color:#2ecc71;font-weight:bold;font-size:11px;">🟩 Laser — Fundo dos Poços</div>
</div>

<div id="toast"></div>
<div id="container-wrapper"><canvas id="gel-canvas"></canvas></div>

<script>
const S=window.Streamlit;
const container=document.getElementById('container-wrapper');
const canvas=document.getElementById('gel-canvas');
const ctx=canvas.getContext('2d');
const img=new Image();

let originalW=0, originalH=0, scale=1;
let yGuia=0, bandas=[], list1=[], list2=[], w1=0, w2=0, skip1=0, skip2=0;
let edges1=[], edges2=[], isReady=false;

let isEditingLanes=false;
let draggingEdge=null;
let hoverEdge=null;

// PAN (mover)
let isPanning=false;
let panStartX=0, panStartY=0, scrollStartL=0, scrollStartT=0;
let spacePressed=false;

function toast(msg,type='success'){
    const t=document.getElementById('toast');
    const c={success:'rgba(46,204,113,.96)',info:'rgba(52,152,219,.96)',warning:'rgba(241,196,15,.96)',error:'rgba(231,76,60,.96)'};
    t.style.background=c[type]||c.success; t.textContent=msg; t.style.display='block';
    setTimeout(()=>{t.style.display='none';},2200);
}
function instrucao(html){ document.getElementById('calib-instrucao').innerHTML=html; }
function sendData(){ if(S) S.setComponentValue({bandas, calib_edges1:edges1, calib_edges2:edges2}); }

function clearLanes(){
    if(confirm('Deseja APAGAR todas as linhas e desenhar do zero?')){
        edges1=[]; edges2=[]; draw(); sendData();
        toast('Grade apagada!','info');
        if(!isEditingLanes) toggleEditLanes();
    }
}

function toggleEditLanes(){
    isEditingLanes=!isEditingLanes;
    const btn=document.getElementById('btn-edit');
    if(isEditingLanes){
        btn.className='btn-edit ativo';
        btn.textContent='✅ Salvar Colunas e Voltar';
        instrucao('<strong>MODO COLUNAS:</strong> Esq=criar/arrastar linha · Dir=apagar linha · Meio/Espaço=mover · Roda=zoom');
        toast('Modo colunas ativo','info');
    } else {
        btn.className='btn-edit';
        btn.textContent='📐 Ajustar / Criar Colunas Manualmente';
        instrucao('<strong>Zoom:</strong> roda · <strong>Mover:</strong> botão do meio OU Espaço+arrastar · <strong>Colunas:</strong> ative o modo azul');
        toast('Colunas salvas!','success');
        sendData();
    }
    draw();
}

function getNearestEdge(mX){
    let minDist=8/scale, nearest=null;
    if(edges1){
        for(let i=0;i<edges1.length;i++){
            const d=Math.abs(mX-edges1[i]);
            if(d<minDist){ minDist=d; nearest={group:1,idx:i}; }
        }
    }
    if(edges2){
        for(let i=0;i<edges2.length;i++){
            const d=Math.abs((mX-w1)-edges2[i]);
            if(d<minDist){ minDist=d; nearest={group:2,idx:i}; }
        }
    }
    return nearest;
}

function laneFromEdges(edges,x){
    if(!edges||edges.length<2) return -1;
    for(let k=0;k<edges.length-1;k++){
        if(x>=edges[k]&&x<edges[k+1]) return k;
    }
    if(x>=edges[edges.length-1]) return edges.length-2;
    return -1;
}

function draw(){
    if(!img.complete||originalW===0) return;
    canvas.width=Math.round(originalW*scale);
    canvas.height=Math.round(originalH*scale);
    ctx.imageSmoothingEnabled=true; ctx.imageSmoothingQuality='high';
    ctx.clearRect(0,0,canvas.width,canvas.height);
    ctx.drawImage(img,0,0,canvas.width,canvas.height);

    const xSplit=w1*scale;
    const e1=edges1?edges1.map(v=>v*scale):[];
    const e2=edges2?edges2.map(v=>xSplit+v*scale):[];

    function drawEdgeLine(xCanvas,isId,isHovered){
        ctx.save(); ctx.beginPath();
        ctx.moveTo(xCanvas,0); ctx.lineTo(xCanvas,canvas.height);
        if(isEditingLanes){
            ctx.strokeStyle=isHovered?'#e74c3c':'#3498db';
            ctx.lineWidth=isHovered?Math.max(2.5,3*scale):Math.max(1.5,1.5*scale);
            ctx.setLineDash([8,4]);
        } else {
            ctx.strokeStyle=isId?'rgba(255,220,50,0.7)':'rgba(255,255,255,0.15)';
            ctx.lineWidth=isId?Math.max(1,1.1*scale):Math.max(0.4,0.6*scale);
        }
        ctx.stroke(); ctx.restore();
    }

    for(let k=0;k<e1.length;k++){
        if(e1[k]<-1||e1[k]>xSplit+1) continue;
        const isId=(k>=skip1&&k<skip1+list1.length);
        const isHov=hoverEdge&&hoverEdge.group===1&&hoverEdge.idx===k;
        drawEdgeLine(e1[k],isId,isHov);
    }
    for(let k=0;k<e2.length;k++){
        if(e2[k]<xSplit-1||e2[k]>canvas.width+1) continue;
        const isId=(k>=skip2&&k<skip2+list2.length);
        const isHov=hoverEdge&&hoverEdge.group===2&&hoverEdge.idx===k;
        drawEdgeLine(e2[k],isId,isHov);
    }

    if(yGuia>0){
        const yL=yGuia*scale;
        ctx.save(); ctx.beginPath();
        ctx.moveTo(0,yL); ctx.lineTo(canvas.width,yL);
        ctx.strokeStyle='#2ecc71'; ctx.lineWidth=Math.max(1,1.5*scale);
        ctx.setLineDash([10,5]); ctx.stroke(); ctx.setLineDash([]); ctx.restore();
    }

    const yLaserPx=yGuia>0?yGuia*scale:canvas.height*0.15;
    const yId=Math.max(12,yLaserPx*0.55);

    function drawID(txt,cx,y,fs){
        ctx.save();
        ctx.font=`bold ${fs}px Arial`;
        ctx.textAlign='center'; ctx.textBaseline='middle';
        ctx.lineWidth=Math.max(2,2.5*scale);
        ctx.strokeStyle='rgba(0,0,0,0.95)'; ctx.strokeText(txt,cx,y);
        ctx.fillStyle='#ffffff'; ctx.fillText(txt,cx,y);
        ctx.restore();
    }

    if(e1.length>1){
        for(let i=0;i<list1.length;i++){
            const lane=i+skip1;
            if(lane<0||lane>=e1.length-1) continue;
            const cx=(e1[lane]+e1[lane+1])/2;
            const fs=Math.max(7,Math.min(13,Math.round((e1[lane+1]-e1[lane])*0.38)));
            if(cx<0||cx>xSplit) continue;
            drawID(list1[i],cx,yId,fs);
        }
    }
    if(e2.length>1){
        for(let i=0;i<list2.length;i++){
            const lane=i+skip2;
            if(lane<0||lane>=e2.length-1) continue;
            const cx=(e2[lane]+e2[lane+1])/2;
            const fs=Math.max(7,Math.min(13,Math.round((e2[lane+1]-e2[lane])*0.38)));
            if(cx<xSplit||cx>canvas.width) continue;
            drawID(list2[i],cx,yId,fs);
        }
    }

    bandas.forEach((b,index)=>{
        const yD=b.y*scale;
        const letra=index<26?String.fromCharCode(97+index):('b'+index);
        ctx.save(); ctx.beginPath(); ctx.moveTo(0,yD); ctx.lineTo(canvas.width,yD);
        ctx.strokeStyle='rgba(255,255,255,0.30)'; ctx.lineWidth=Math.max(0.8,1*scale);
        ctx.setLineDash([5,4]); ctx.stroke(); ctx.setLineDash([]); ctx.restore();

        const fB=Math.max(9,Math.min(14,Math.round(12*scale)));
        ctx.save(); ctx.font=`bold ${fB}px Arial`; ctx.textBaseline='bottom';
        ctx.lineWidth=3; ctx.strokeStyle='rgba(0,0,0,0.85)'; ctx.fillStyle='#ffff00';
        ctx.textAlign='left'; ctx.strokeText(letra,5,yD-2); ctx.fillText(letra,5,yD-2);
        ctx.textAlign='right'; ctx.strokeText(letra,canvas.width-5,yD-2); ctx.fillText(letra,canvas.width-5,yD-2);
        ctx.restore();

        if(b.marks){
            b.marks.forEach((mark,colIdx)=>{
                if(mark!==1) return;
                let cx=-1,tw=20*scale;
                if(colIdx<list1.length){
                    const lane=colIdx+skip1;
                    if(e1&&lane>=0&&lane<e1.length-1){
                        cx=(e1[lane]+e1[lane+1])/2;
                        if(cx<0||cx>xSplit) cx=-1;
                        else tw=Math.max(6,(e1[lane+1]-e1[lane])*0.75);
                    }
                } else {
                    const idx2=colIdx-list1.length, lane=idx2+skip2;
                    if(e2&&lane>=0&&lane<e2.length-1){
                        cx=(e2[lane]+e2[lane+1])/2;
                        if(cx<xSplit||cx>canvas.width) cx=-1;
                        else tw=Math.max(6,(e2[lane+1]-e2[lane])*0.75);
                    }
                }
                if(cx>=0){
                    const th=Math.max(3,4*scale);
                    ctx.save(); ctx.fillStyle='#00ff00';
                    ctx.fillRect(cx-tw/2,yD-th/2,tw,th);
                    ctx.strokeStyle='#004400'; ctx.lineWidth=Math.max(0.5,0.7*scale);
                    ctx.strokeRect(cx-tw/2,yD-th/2,tw,th); ctx.restore();
                }
            });
        }
    });
}

function onRender(ev){
    const payload=(ev&&ev.detail)?ev.detail:ev;
    const args=payload&&payload.args?payload.args:null;
    if(!args) return;

    w1=args.w1||0; w2=args.w2||0; yGuia=args.y_guia||0;
    list1=args.list1||[]; list2=args.list2||[];
    skip1=args.skip1||0; skip2=args.skip2||0;

    if(!isReady){
        edges1=args.edges1||[];
        edges2=args.edges2||[];
        if(args.bandas_init) bandas=args.bandas_init;
    } else {
        if(args.force_edges){
            edges1=args.edges1||[];
            edges2=args.edges2||[];
        }
    }

    const src="data:image/jpeg;base64,"+args.img_b64;
    if(img.src!==src){
        img.src=src;
        img.onload=()=>{
            originalW=img.width; originalH=img.height;
            scale=args.largura_inicial/originalW;
            draw(); if(S) S.setFrameHeight(700);
        };
    } else { draw(); }
    isReady=true;
}

if(S){ S.events.addEventListener("streamlit:render",onRender); S.setComponentReady(); }

// ── ZOOM (roda do mouse) ─────────────────────────────────────────────────────
container.addEventListener('wheel',(e)=>{
    e.preventDefault();
    const rect=canvas.getBoundingClientRect();
    const mx=e.clientX-rect.left, my=e.clientY-rect.top;
    const zf=e.deltaY<0?1.15:0.85;
    const ns=scale*zf;
    if(ns*originalW<400||ns*originalW>14000) return;
    scale=ns; draw();
    const cr=container.getBoundingClientRect();
    container.scrollLeft=mx*zf-(e.clientX-cr.left);
    container.scrollTop=my*zf-(e.clientY-cr.top);
},{passive:false});

document.addEventListener('keydown',(e)=>{
    if(e.code==='Space' && !e.repeat){
        spacePressed=true;
        if(!isEditingLanes || !draggingEdge) canvas.style.cursor='grab';
        e.preventDefault();
    }
    if(e.key==='Escape'){
        if(isEditingLanes){ toggleEditLanes(); return; }
        if(bandas.length>0&&confirm('🗑️ Limpar todas as bandas?')){
            bandas=[]; draw(); sendData(); toast('✓ Bandas removidas','info');
        }
    }
});
document.addEventListener('keyup',(e)=>{
    if(e.code==='Space'){
        spacePressed=false;
        if(!isPanning) canvas.style.cursor=isEditingLanes?(hoverEdge?'col-resize':'crosshair'):'crosshair';
    }
});

canvas.addEventListener('mousemove',(e)=>{
    const rect=canvas.getBoundingClientRect();
    const mX=(e.clientX-rect.left)/scale;

    if(isPanning){
        container.scrollLeft=scrollStartL-(e.clientX-panStartX);
        container.scrollTop=scrollStartT-(e.clientY-panStartY);
        return;
    }

    if(isEditingLanes){
        if(draggingEdge){
            if(draggingEdge.group===1) edges1[draggingEdge.idx]=mX;
            else edges2[draggingEdge.idx]=mX-w1;
            draw();
        } else {
            const oldH=hoverEdge?`${hoverEdge.group}-${hoverEdge.idx}`:null;
            hoverEdge=getNearestEdge(mX);
            const newH=hoverEdge?`${hoverEdge.group}-${hoverEdge.idx}`:null;
            if(!spacePressed) canvas.style.cursor=hoverEdge?'col-resize':'crosshair';
            if(oldH!==newH) draw();
        }
    }
});

canvas.addEventListener('mousedown',(e)=>{
    const rect=canvas.getBoundingClientRect();
    const mX=(e.clientX-rect.left)/scale;
    const mY=(e.clientY-rect.top)/scale;
    const totalInd=list1.length+list2.length;

    if(e.button===1 || (e.button===0 && spacePressed)){
        e.preventDefault();
        isPanning=true;
        panStartX=e.clientX; panStartY=e.clientY;
        scrollStartL=container.scrollLeft; scrollStartT=container.scrollTop;
        canvas.style.cursor='grabbing';
        return;
    }

    if(isEditingLanes){
        e.preventDefault(); e.stopPropagation();
        const nearest=getNearestEdge(mX);

        if(e.button===0){
            if(nearest){
                draggingEdge=nearest;
            } else {
                if(mX<=w1){
                    if(!edges1) edges1=[];
                    edges1.push(mX); edges1.sort((a,b)=>a-b);
                } else {
                    if(!edges2) edges2=[];
                    edges2.push(mX-w1); edges2.sort((a,b)=>a-b);
                }
                draw(); sendData();
            }
        } else if(e.button===2){
            if(nearest){
                if(nearest.group===1) edges1.splice(nearest.idx,1);
                else edges2.splice(nearest.idx,1);
                hoverEdge=null; draw(); sendData();
            }
        }
        return;
    }

    if(e.button===0){
        let clickedB=null;
        for(let b of bandas){ if(Math.abs(b.y-mY)<18/scale){clickedB=b;break;} }

        if(clickedB){
            let col=-1;
            if(mX<w1){
                const lane=laneFromEdges(edges1,mX);
                if(lane===-1){toast('Fora das colunas','warning');return;}
                const idx=lane-skip1;
                if(idx<0||idx>=list1.length){toast('Coluna ignorada (ladder/vazia)','info');return;}
                col=idx;
            } else {
                const lane=laneFromEdges(edges2,mX-w1);
                if(lane===-1){toast('Fora das colunas','warning');return;}
                const idx2=lane-skip2;
                if(idx2<0||idx2>=list2.length){toast('Coluna ignorada (ladder/vazia)','info');return;}
                col=list1.length+idx2;
            }
            if(col>=0&&col<totalInd){
                if(!clickedB.marks) clickedB.marks=new Array(totalInd).fill(0);
                if(clickedB.marks.length<totalInd){
                    clickedB.marks=clickedB.marks.concat(new Array(totalInd-clickedB.marks.length).fill(0));
                }
                clickedB.marks[col]=clickedB.marks[col]===1?0:1;
                const nome=col<list1.length?list1[col]:list2[col-list1.length];
                toast((clickedB.marks[col]?'✓ Marcado: ':'✗ Desmarcado: ')+nome, clickedB.marks[col]?'success':'warning');
                draw(); sendData();
            }
        } else {
            bandas.push({y:mY,marks:new Array(totalInd).fill(0)});
            bandas.sort((a,b)=>a.y-b.y);
            const li=bandas.findIndex(b=>b.y===mY);
            const letra=li<26?String.fromCharCode(97+li):`b${li}`;
            toast(`✓ Linha "${letra}" criada!`,'info');
            draw(); sendData();
        }
    } else if(e.button===2){
        if(bandas.length===0) return;
        let ci=0,md=Math.abs(bandas[0].y-mY);
        for(let i=1;i<bandas.length;i++){const d=Math.abs(bandas[i].y-mY);if(d<md){md=d;ci=i;}}
        if(md<25/scale){
            const le=ci<26?String.fromCharCode(97+ci):`b${ci}`;
            bandas.splice(ci,1); toast(`✗ Linha "${le}" excluída!`,'warning');
            draw(); sendData();
        }
    }
});

function endPanOrDrag(){
    if(isPanning){
        isPanning=false;
        canvas.style.cursor=spacePressed?'grab':(isEditingLanes?(hoverEdge?'col-resize':'crosshair'):'crosshair');
    }
    if(draggingEdge){
        if(edges1) edges1.sort((a,b)=>a-b);
        if(edges2) edges2.sort((a,b)=>a-b);
        draggingEdge=null;
        draw(); sendData();
    }
}
canvas.addEventListener('mouseup',endPanOrDrag);
canvas.addEventListener('mouseleave',endPanOrDrag);
canvas.addEventListener('contextmenu',e=>e.preventDefault());
canvas.addEventListener('auxclick',e=>{ if(e.button===1) e.preventDefault(); });
</script>
</body>
</html>
"""

COMP_VERSION = "v28_2_" + hashlib.md5(HTML_TEMPLATE.encode("utf-8")).hexdigest()[:10]

@st.cache_resource
def get_interactive_canvas(version=COMP_VERSION):
    comp_dir = os.path.join(tempfile.gettempdir(), f"ssr_canvas_{version}")
    os.makedirs(comp_dir, exist_ok=True)
    fp = os.path.join(comp_dir, "index.html")
    with open(fp, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE)
    return components.declare_component(f"ssr_canvas_{version}", path=comp_dir)


# ============================================================
#  ESTADO DA SESSÃO
# ============================================================
defaults = {
    "offset_y2":0, "rot_f1":0.0, "rot_f2":0.0, "escala_y2":1.0,
    "pos_laser":100, "auto_ok":False, "y1_det":None, "y2_det":None,
    "last_upload_id":None, "limpeza_nivel":"Desligado",
    "skip1":1, "skip2":0, "auto_update":True,
    "ordenar_ids_auto": True,
    "todas_matrizes": {},
    "f1_select": [],
    "f2_select": []
}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

# Exibição de mensagem pendente de importação sem congelar Python
if "msg_import_pendente" in st.session_state:
    st.success(st.session_state["msg_import_pendente"])
    del st.session_state["msg_import_pendente"]

# --- TENTA RECUPERAR BACKUP DO DISCO NA PRIMEIRA EXECUÇÃO ---
if "backup_tentado" not in st.session_state:
    recuperado, total_primers = carregar_backup_local()
    st.session_state["backup_tentado"] = True
    if recuperado and total_primers > 0:
        st.toast(f"🔄 Backup recuperado do disco: {total_primers} primers recarregados!", icon="💾")

# ============================================================
#  INTERFACE DO APLICATIVO
# ============================================================
st.title("🧬 Sistema SSR — Leitor de Gel Duplo v28.5")
st.caption("🚀 v28.5: Salvamento de Géis de Conferência (JPG) · Auto-Criar Pastas no PC · Importação com Nomes")

pasta_bkp_abs, pasta_geis_abs = garantir_pastas()
with st.expander("📁 Exibir caminhos de salvamento neste computador"):
    st.write(f"**Backup e Planilhas:** `{pasta_bkp_abs}`")
    st.write(f"**Fotos dos Géis com Marcações:** `{pasta_geis_abs}`")

aba1,aba2,aba3,aba4 = st.tabs(["📋 1. Calibração e Marcação","📊 2. Dendrograma UPGMA","📥 3. Exportar / Importar Excel","❓ 4. Ajuda"])

with aba1:
    st.header("⚙️ Configuração das Amostras")
    
    opcoes_ids = [str(i) for i in range(1, 301)] + ["L", "C", "C1", "C2", "M"]

    c1, c2, c3 = st.columns([2,4,4])
    with c1:
        nome_primer = st.text_input("Nome do Primer:", value="ISSR 19")
        st.session_state["auto_update"] = st.checkbox(
            "⚡ Atualização automática",
            value=bool(st.session_state.get("auto_update",True)),
            help="Quando ligado, mudanças no canvas (bandas/colunas) atualizam a matriz na hora."
        )
        st.session_state["ordenar_ids_auto"] = st.checkbox(
            "🔢 Ordenar IDs (crescente)",
            value=bool(st.session_state.get("ordenar_ids_auto", True)),
            help="Quando marcado, organiza os números de 1 a 300 em ordem crescente. Quando desmarcado, mantém a ordem exata em que você clicou."
        )

    with c2:
        lista_f1_raw = st.multiselect(
            "Indivíduos Foto 1 (Esquerda):",
            options=opcoes_ids,
            key="f1_select",
            on_change=sync_e_ordenar_f1,
            help="Selecione os números."
        )
        
        if st.session_state["ordenar_ids_auto"]:
            lista_f1 = ordenar_ids(st.session_state["f1_select"])
        else:
            lista_f1 = list(dict.fromkeys(st.session_state["f1_select"]))
            
        if not lista_f1:
            st.error("❌ Lista vazia")
        else:
            status_txt = "ordenados" if st.session_state["ordenar_ids_auto"] else "ordem da escolha"
            st.success(f"✅ {len(lista_f1)} indivíduos selecionados ({status_txt})")
            
    with c3:
        lista_f2_raw = st.multiselect(
            "Indivíduos Foto 2 (Direita):",
            options=opcoes_ids,
            key="f2_select",
            on_change=sync_e_ordenar_f2,
            help="Selecione os números."
        )
        
        if st.session_state["ordenar_ids_auto"]:
            lista_f2 = ordenar_ids(st.session_state["f2_select"])
        else:
            lista_f2 = list(dict.fromkeys(st.session_state["f2_select"]))
            
        if not lista_f2:
            st.error("❌ Lista vazia")
        else:
            status_txt = "ordenados" if st.session_state["ordenar_ids_auto"] else "ordem da escolha"
            st.success(f"✅ {len(lista_f2)} indivíduos selecionados ({status_txt})")

    if lista_f1 and lista_f2:
        lista_uni = list(dict.fromkeys(lista_f1 + lista_f2))
        st.info(f"📊 **Total: {len(lista_uni)} indivíduos únicos**")
    else:
        lista_uni = []

    st.divider()
    st.header("📸 Upload das Fotos")
    cu1,cu2=st.columns(2)
    with cu1: foto1=st.file_uploader("FOTO 1 (Esquerda)",type=["png","jpg","jpeg","tif"])
    with cu2: foto2=st.file_uploader("FOTO 2 (Direita)",type=["png","jpg","jpeg","tif"])

    if foto1 and foto2 and lista_f1 and lista_f2:
        img1,er1=carregar_imagem(foto1); img2,er2=carregar_imagem(foto2)
        if img1 is None: st.error(f"❌ Foto 1: {er1}"); st.stop()
        if img2 is None: st.error(f"❌ Foto 2: {er2}"); st.stop()
        upload_id=f"{foto1.name}_{foto1.size}_{foto2.name}_{foto2.size}"

        st.markdown("---"); st.subheader("🧹 Limpeza de Sujeira do Gel")
        cl1,cl2,cl3=st.columns([2,2,3])
        with cl1:
            st.session_state["limpeza_nivel"]=st.selectbox(
                "Intensidade:",["Desligado","Leve","Médio","Forte"],
                index=["Desligado","Leve","Médio","Forte"].index(st.session_state["limpeza_nivel"]))
        with cl2: aplicar_antes_cor=st.checkbox("Aplicar antes do filtro de cor",value=True)
        with cl3:
            if st.session_state["limpeza_nivel"]!="Desligado":
                st.info(f"🧹 Modo **{st.session_state['limpeza_nivel']}** ativado")

        img1_work=remover_sujeira(img1,st.session_state["limpeza_nivel"]) if aplicar_antes_cor else img1.copy()
        img2_work=remover_sujeira(img2,st.session_state["limpeza_nivel"]) if aplicar_antes_cor else img2.copy()

        st.markdown("---"); st.subheader("🛠️ Calibração Vertical")
        colA,colB=st.columns([2,3])
        with colA:
            if st.button("🎯 ALINHAR AUTOMATICAMENTE",type="primary",use_container_width=True):
                off,ylaser,y1d,y2d=auto_calibrar(img1_work,img2_work,
                    rot1=st.session_state["rot_f1"],rot2=st.session_state["rot_f2"],escala2=st.session_state["escala_y2"])
                st.session_state.update({"offset_y2":off,"pos_laser":ylaser,"y1_det":y1d,"y2_det":y2d,"auto_ok":True,"last_upload_id":upload_id})
                st.rerun()
            if st.session_state["last_upload_id"]!=upload_id:
                off,ylaser,y1d,y2d=auto_calibrar(img1_work,img2_work)
                st.session_state.update({"offset_y2":off,"pos_laser":ylaser,"y1_det":y1d,"y2_det":y2d,"auto_ok":True,"last_upload_id":upload_id})
                st.rerun()
        with colB:
            if st.session_state["auto_ok"]:
                st.success(f"✅ F1 Y={st.session_state['y1_det']} | F2 Y={st.session_state['y2_det']} | Offset={st.session_state['offset_y2']}px")

        cc1,cc2,cc3=st.columns([2,3,3])
        with cc1:
            st.markdown("🎯 **Ajuste Fino Vertical**")
            b1,b2,b3,b4=st.columns(4)
            if b1.button("⬆️+10"): st.session_state["offset_y2"]-=10; st.rerun()
            if b2.button("⬆️+1"):  st.session_state["offset_y2"]-=1;  st.rerun()
            if b3.button("⬇️-1"):  st.session_state["offset_y2"]+=1;  st.rerun()
            if b4.button("⬇️-10"): st.session_state["offset_y2"]+=10; st.rerun()
            st.session_state["offset_y2"]=st.number_input("Deslocamento Y:",value=int(st.session_state["offset_y2"]),step=1)
        with cc2:
            st.markdown("📏 **Laser + Filtro**")
            st.session_state["pos_laser"]=st.slider("Altura Linha Guia:",5,500,int(st.session_state["pos_laser"]),1)
            filtro_cor=st.selectbox("Filtro de Cor:",["Original","Preto e Branco (Fundo Preto)","Preto e Branco (Invertido - Fundo Branco)"])
        with cc3:
            st.markdown("🔄 **Rotação**")
            st.session_state["rot_f1"]=st.slider("Rotação F1 (°):",-10.0,10.0,float(st.session_state["rot_f1"]),0.1)
            st.session_state["rot_f2"]=st.slider("Rotação F2 (°):",-10.0,10.0,float(st.session_state["rot_f2"]),0.1)
            if st.button("🔁 Recalibrar",use_container_width=True):
                off,ylaser,y1d,y2d=auto_calibrar(img1_work,img2_work,rot1=st.session_state["rot_f1"],rot2=st.session_state["rot_f2"],escala2=st.session_state["escala_y2"])
                st.session_state.update({"offset_y2":off,"pos_laser":ylaser,"auto_ok":True}); st.rerun()

        with st.expander("🔍 Ajustes Adicionais"):
            ce1,ce2=st.columns(2)
            with ce1:
                brilho=st.slider("Brilho:",0.5,3.0,1.0,0.1)
                contraste=st.slider("Contraste:",0.5,3.0,1.0,0.1)
            with ce2:
                st.session_state["escala_y2"]=st.slider("Escala F2 (%):",80.0,120.0,float(st.session_state["escala_y2"]*100),0.5)/100.0

        cor_fundo=(255,255,255) if "Invertido" in filtro_cor else (0,0,0)
        img1_p=aplicar_filtro_bw(img1_work,filtro_cor,brilho,contraste)
        img2_p=aplicar_filtro_bw(img2_work,filtro_cor,brilho,contraste)
        if not aplicar_antes_cor:
            img1_p=remover_sujeira(img1_p,st.session_state["limpeza_nivel"])
            img2_p=remover_sujeira(img2_p,st.session_state["limpeza_nivel"])
        img1_t=transformar_imagem(img1_p,st.session_state["rot_f1"],1.0,0,cor_fundo)
        img2_t=transformar_imagem(img2_p,st.session_state["rot_f2"],st.session_state["escala_y2"],st.session_state["offset_y2"],cor_fundo)
        h1,w1=img1_t.shape[:2]; h2,w2=img2_t.shape[:2]
        if h2!=h1:
            nw2=int(w2*(h1/h2)); img2_t=cv2.resize(img2_t,(nw2,h1),interpolation=cv2.INTER_AREA)
        cv2.line(img1_t,(w1-2,0),(w1-2,h1),(255,150,0),3)
        cv2.line(img2_t,(1,0),(1,h1),(255,150,0),3)
        img_unida=cv2.hconcat([img1_t,img2_t])
        img_rgb=cv2.cvtColor(img_unida,cv2.COLOR_BGR2RGB)
        _,buffer=cv2.imencode(".jpg",img_rgb,[cv2.IMWRITE_JPEG_QUALITY,96])
        img_b64=base64.b64encode(buffer).decode()

        with st.expander("🧭 Pular colunas iniciais (ladder/controle)"):
            col_s1,col_s2=st.columns(2)
            with col_s1: st.session_state["skip1"]=st.number_input("F1 — colunas a pular",0,10,int(st.session_state["skip1"]),1)
            with col_s2: st.session_state["skip2"]=st.number_input("F2 — colunas a pular",0,10,int(st.session_state["skip2"]),1)

        st.divider(); st.subheader("🔬 Visualizador Interativo")
        
        largura_panoramica=st.slider("Tamanho Inicial (px):",1000,6000,2600,100)

        cols1=len(lista_f1)+int(st.session_state["skip1"])
        cols2=len(lista_f2)+int(st.session_state["skip2"])
        canvas_key=f"canvas_{upload_id}"

        if f"edges1_{canvas_key}" in st.session_state:
            edges1_init=st.session_state[f"edges1_{canvas_key}"]
        else:
            edges1_init=np.linspace(0,w1,cols1+1).round().astype(int).tolist()

        if f"edges2_{canvas_key}" in st.session_state:
            edges2_init=st.session_state[f"edges2_{canvas_key}"]
        else:
            edges2_init=np.linspace(0,img2_t.shape[1],cols2+1).round().astype(int).tolist()

        interactive_canvas=get_interactive_canvas()
        canvas_result=interactive_canvas(
            img_b64=img_b64, w1=int(w1), w2=int(img2_t.shape[1]),
            y_guia=int(st.session_state["pos_laser"]),
            list1=lista_f1, list2=lista_f2,
            largura_inicial=int(largura_panoramica),
            skip1=int(st.session_state["skip1"]), skip2=int(st.session_state["skip2"]),
            edges1=edges1_init, edges2=edges2_init,
            bandas_init=st.session_state.get(f"bandas_{canvas_key}",[]),
            key=canvas_key
        )

        auto_upd=bool(st.session_state.get("auto_update",True))
        if canvas_result and isinstance(canvas_result,dict):
            changed=False
            if "bandas" in canvas_result and canvas_result["bandas"]!=st.session_state.get(f"bandas_{canvas_key}",[]):
                st.session_state[f"bandas_{canvas_key}"]=canvas_result["bandas"]; changed=True
            e1_res=canvas_result.get("calib_edges1")
            e2_res=canvas_result.get("calib_edges2")
            if e1_res is not None and e1_res!=st.session_state.get(f"edges1_{canvas_key}"):
                st.session_state[f"edges1_{canvas_key}"]=e1_res; changed=True
            if e2_res is not None and e2_res!=st.session_state.get(f"edges2_{canvas_key}"):
                st.session_state[f"edges2_{canvas_key}"]=e2_res; changed=True
            if changed and auto_upd:
                st.rerun()

        bandas_salvas=st.session_state.get(f"bandas_{canvas_key}",[])
        num_canvas_bands=len(bandas_salvas)

        st.divider(); st.subheader("✍️ Matriz de Leitura")
        cm1,cm2=st.columns([2,4])
        with cm1: n_bandas=st.number_input("Total de bandas:",min_value=1,max_value=100,value=max(1,num_canvas_bands))
        n_real=max(n_bandas,num_canvas_bands)
        with cm2: st.info(f"📊 **{n_real} bandas** × **{len(lista_uni)} indivíduos**")

        letras_tabela=[chr(97+i) if i<26 else f"b{i}" for i in range(n_real)]
        key_mat=f"matriz_dados_{nome_primer}_{upload_id}"
        dados_bool={ind:[False]*n_real for ind in lista_uni}
        for i,b in enumerate(bandas_salvas):
            if i>=n_real: break
            for col_idx,m in enumerate(b.get("marks",[])):
                if m==1:
                    if col_idx<len(lista_f1): ind_name=lista_f1[col_idx]
                    else:
                        idx2=col_idx-len(lista_f1)
                        ind_name=lista_f2[idx2] if idx2<len(lista_f2) else None
                    if ind_name and ind_name in dados_bool: dados_bool[ind_name][i]=True

        df_at=pd.DataFrame(dados_bool,index=letras_tabela); df_at.index.name="Banda"
        st.session_state[key_mat]=df_at
        df_ed=st.data_editor(st.session_state[key_mat],use_container_width=True,height=350,key=f"editor_{key_mat}")
        
        st.session_state["matriz_final"]=df_ed
        st.session_state["primer_nome"]=nome_primer

        tm=df_ed.sum().sum(); tc=n_real*len(lista_uni); taxa=(tm/tc*100) if tc>0 else 0
        cs1,cs2,cs3=st.columns(3)
        cs1.metric("Total Marcado",f"{int(tm)}/{tc}")
        cs2.metric("Taxa de Presença",f"{taxa:.1f}%")
        cs3.metric("Bandas × Indivíduos",f"{n_real} × {len(lista_uni)}")

        # ---------------------------------------------------------
        # GERENCIADOR DE PRIMERS (SALVA MATRIZ + FOTO DO GEL)
        # ---------------------------------------------------------
        st.divider()
        st.subheader("💾 Gerenciador de Primers (Multilocus)")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown(f"**Salvar Primer Atual:**\nClique abaixo para gravar a matriz Excel e salvar o **Gel Processado com Marcações** no PC.")
            if st.button(f"💾 Salvar matriz e imagem do primer: {nome_primer}", type="primary", use_container_width=True):
                df_to_save = df_ed.copy()
                df_to_save.index = [f"{nome_primer}_{x}" for x in df_to_save.index]
                st.session_state["todas_matrizes"][nome_primer] = df_to_save
                
                # 1. Salva Backup e Excel
                sucesso_backup, desc_caminho = salvar_backup_local()
                
                # 2. Salva Foto do Gel com as marcações
                e1_now = st.session_state.get(f"edges1_{canvas_key}", edges1_init)
                e2_now = st.session_state.get(f"edges2_{canvas_key}", edges2_init)
                ok_gel, info_gel = salvar_gel_conferencia(
                    img_bgr=img_unida,
                    nome_primer=nome_primer,
                    bandas_salvas=bandas_salvas,
                    lista_f1=lista_f1,
                    lista_f2=lista_f2,
                    edges1=e1_now,
                    edges2=e2_now,
                    w1=int(w1),
                    skip1=int(st.session_state["skip1"]),
                    skip2=int(st.session_state["skip2"]),
                    y_guia=int(st.session_state["pos_laser"])
                )
                
                if sucesso_backup:
                    msg = f"✅ Primer **{nome_primer}** gravado no disco!\n\n📁 **Matriz/Excel:** `{desc_caminho}`"
                    if ok_gel:
                        msg += f"\n\n🖼️ **Gel Processado:** `{info_gel}`"
                    st.success(msg)
                else:
                    st.error(f"❌ Erro ao gerar backup físico em disco: {desc_caminho}")

            if st.button("🖼️ Salvar apenas a foto do GEL com marcações", use_container_width=True):
                e1_now = st.session_state.get(f"edges1_{canvas_key}", edges1_init)
                e2_now = st.session_state.get(f"edges2_{canvas_key}", edges2_init)
                ok_gel, info_gel = salvar_gel_conferencia(
                    img_bgr=img_unida,
                    nome_primer=nome_primer,
                    bandas_salvas=bandas_salvas,
                    lista_f1=lista_f1,
                    lista_f2=lista_f2,
                    edges1=e1_now,
                    edges2=e2_now,
                    w1=int(w1),
                    skip1=int(st.session_state["skip1"]),
                    skip2=int(st.session_state["skip2"]),
                    y_guia=int(st.session_state["pos_laser"])
                )
                if ok_gel:
                    st.success(f"🖼️ Gel gravado em:\n`{info_gel}`")
                else:
                    st.error(f"❌ Não foi possível salvar o gel: {info_gel}")

        with col_m2:
            st.markdown("**📂 Continuar a partir de Planilha Pronta:**\nImporte um Excel gerado anteriormente pelo App para continuar editando de onde parou:")
            arq_excel_a1 = st.file_uploader("Selecionar planilha Excel para carregar dados:", type=["xlsx"], key="importar_excel_a1")
            if arq_excel_a1:
                try:
                    dict_primers, erro = importar_excel_completo(arq_excel_a1.read())
                    if erro:
                        st.error(f"❌ Falha ao ler arquivo: {erro}")
                    else:
                        st.session_state["todas_matrizes"].update(dict_primers)
                        salvar_backup_local()
                        nomes_lidos = ", ".join(list(dict_primers.keys()))
                        st.session_state["msg_import_pendente"] = f"✅ Planilha restaurada! {len(dict_primers)} primers carregados:\n\n**{nomes_lidos}**"
                        st.rerun()
                except Exception as ex:
                    st.error(f"❌ Falha crítica ao importar: {str(ex)}")

        st.divider()
        salvos = list(st.session_state["todas_matrizes"].keys())
        if salvos:
            st.info(f"📦 **{len(salvos)} Primers carregados em cache:** {', '.join(salvos)}")
            if st.button("🗑️ Limpar toda a memória de primers"):
                st.session_state["todas_matrizes"] = {}
                if os.path.exists(BACKUP_FILE):
                    try:
                        os.remove(BACKUP_FILE)
                    except:
                        pass
                st.rerun()
        else:
            st.info("Nenhum primer na memória temporária.")

    elif not(lista_f1 and lista_f2): 
        st.info("💡 Selecione os indivíduos de cada foto (caixas acima) para poder enviar as imagens.")
    else: 
        st.info("💡 Faça upload das duas fotos para iniciar.")


with aba2:
    st.header("📊 Análise de Diversidade Genética (UPGMA Combinado)")
    
    salvos = list(st.session_state.get("todas_matrizes", {}).keys())
    
    if len(salvos) == 0:
        st.info("💡 Salve ou Importe um primer na **Aba 1 (Gerenciador de Primers)** para gerar o Dendrograma.")
    else:
        st.markdown("### 🗂️ Quais primers você quer analisar juntos?")
        primers_selecionados = st.multiselect("Selecione os primers:", salvos, default=salvos)
        
        if primers_selecionados:
            dfs_to_combine = [st.session_state["todas_matrizes"][p] for p in primers_selecionados]
            df_combined = pd.concat(dfs_to_combine, axis=0).fillna(False)
            
            acessos = df_combined.columns.tolist()
            bandas_l = df_combined.index.tolist()
            mat_bin = df_combined.values.T.astype(int)
            
            st.info(f"🧬 Analisando um total de **{len(bandas_l)} bandas** unidas de **{len(primers_selecionados)} primer(s)** para **{len(acessos)} indivíduos**.")
            
            if mat_bin.shape[0] >= 3 and np.sum(mat_bin) > 0:
                titulo_upgma = f"UPGMA — {len(primers_selecionados)} Primers Combinados ({len(acessos)} Indivíduos)"
                
                dj = calcular_jaccard(mat_bin)
                Z = fazer_upgma(dj)
                
                fig = plotar_dendrograma(Z, acessos, titulo_upgma)
                st.pyplot(fig)
                plt.close(fig)
                
                cd1, cd2 = st.columns(2)
                nome_arq = "Combinado" if len(primers_selecionados) > 1 else primers_selecionados[0]
                
                with cd1:
                    bf = io.BytesIO()
                    fd = plotar_dendrograma(Z, acessos, titulo_upgma)
                    fd.savefig(bf, format="png", dpi=300, bbox_inches="tight")
                    plt.close(fd)
                    st.download_button("📷 PNG 300dpi", bf.getvalue(), f"dendrograma_{nome_arq}.png", "image/png", use_container_width=True)
                with cd2:
                    bp = io.BytesIO()
                    fp2 = plotar_dendrograma(Z, acessos, titulo_upgma)
                    fp2.savefig(bp, format="pdf", bbox_inches="tight")
                    plt.close(fp2)
                    st.download_button("📄 PDF", bp.getvalue(), f"dendrograma_{nome_arq}.pdf", "application/pdf", use_container_width=True)
                
                st.divider()
                ce1, ce2 = st.columns(2)
                with ce1:
                    st.subheader("🎯 Genitores Contrastantes")
                    aux = dj.copy()
                    np.fill_diagonal(aux, -1)
                    ix = np.unravel_index(np.argmax(aux), aux.shape)
                    st.success(f"🧬 **{acessos[ix[0]]}** × **{acessos[ix[1]]}**\n\n📊 Dissimilaridade: **{aux[ix]:.3f}**")
                    with st.expander("🔝 Top 5"):
                        pares = sorted([(acessos[i], acessos[j], dj[i,j]) for i in range(len(acessos)) for j in range(i+1, len(acessos))], key=lambda x: x[2], reverse=True)
                        for r, (a1, a2, d) in enumerate(pares[:5], 1):
                            st.write(f"**{r}.** {a1} × {a2} → {d:.3f}")
                with ce2:
                    st.subheader("📈 PIC por Banda")
                    ni = mat_bin.shape[0]
                    rows = []
                    for j in range(mat_bin.shape[1]):
                        p = np.sum(mat_bin[:,j]) / ni
                        pv = 1 - p**2 - (1-p)**2
                        rows.append({"Banda": bandas_l[j], "Freq(1)": round(p,3), "PIC": round(pv,3), "Info": "✓" if pv > 0.25 else "✗"})
                    df_pic = pd.DataFrame(rows)
                    st.dataframe(df_pic.style.background_gradient(subset=['PIC'], cmap='RdYlGn'), use_container_width=True, hide_index=True)
                    st.metric("PIC Médio Geral", f"{df_pic['PIC'].mean():.3f}")
                
                with st.expander("📐 Matriz de Jaccard (Combinada)"):
                    df_j = pd.DataFrame(dj, index=acessos, columns=acessos)
                    st.dataframe(df_j.style.background_gradient(cmap="Blues").format("{:.3f}"), use_container_width=True)
                    
            else:
                st.warning("⚠️ Marque pelo menos 3 indivíduos no total para gerar a árvore.")
        else:
            st.warning("Selecione pelo menos um primer para análise.")


with aba3:
    st.header("📥 Exportar / Importar Matrizes SSR")
    
    # CONTAINER DE IMPORTAÇÃO DIRETA NA ABA DE EXCEL
    with st.expander("📂 IMPORTAR DE PLANILHA SSR EXISTENTE (RETOMAR TRABALHO)", expanded=True):
        st.markdown("Se você já possui uma planilha gerada por esta ferramenta no seu computador, faça o upload dela abaixo para restaurar e continuar a trabalhar nela:")
        arq_excel_a3 = st.file_uploader("Fazer upload do arquivo SSR_Combinado.xlsx:", type=["xlsx"], key="importar_excel_a3")
        if arq_excel_a3:
            try:
                dict_primers, erro = importar_excel_completo(arq_excel_a3.read())
                if erro:
                    st.error(f"❌ Não foi possível carregar a planilha: {erro}")
                else:
                    st.session_state["todas_matrizes"].update(dict_primers)
                    salvar_backup_local()
                    nomes_lidos = ", ".join(list(dict_primers.keys()))
                    st.session_state["msg_import_pendente"] = f"✅ Planilha restaurada com sucesso! Primers carregados:\n\n**{nomes_lidos}**"
                    st.rerun()
            except Exception as ex:
                st.error(f"❌ Erro inesperado ao converter planilha: {str(ex)}")

    st.divider()
    st.subheader("📤 Exportar Planilha Excel Atual")
    salvos = list(st.session_state.get("todas_matrizes", {}).keys())
    
    if len(salvos) == 0:
        st.info("💡 Salve pelo menos um primer na **Aba 1 (Gerenciador de Primers)** para poder exportar.")
    else:
        st.markdown("### Selecione os primers para juntar no Excel:")
        primers_export = st.multiselect("Exportar os seguintes primers:", salvos, default=salvos)
        
        if primers_export:
            dict_export = {p: st.session_state["todas_matrizes"][p] for p in primers_export}
            
            dfs = [st.session_state["todas_matrizes"][p] for p in primers_export]
            df_combined = pd.concat(dfs, axis=0).fillna(False)
            
            acessos = df_combined.columns.tolist()
            mat_bin = df_combined.values.T.astype(int)
            
            dj = None
            if mat_bin.shape[0] >= 3 and np.sum(mat_bin) > 0:
                dj = calcular_jaccard(mat_bin)
                
            nome_export = "Combinado" if len(primers_export) > 1 else primers_export[0]
            
            st.write("✅ Matriz Binária Combinada (em blocos por primer) · ✅ Jaccard · ✅ Arial 11 Centrado")
            st.divider()
            
            buf = exportar_excel_completo(dict_export, nome_export, dj, acessos)
            
            cb1, cb2 = st.columns([1, 2])
            with cb1:
                st.download_button("Planilha Gerada - Download Excel Completo", buf, f"SSR_{nome_export}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with cb2:
                st.success(f"✅ Excel gerado com {len(df_combined)} bandas unidas × {len(acessos)} indivíduos únicos.")


with aba4:
    st.header("❓ Guia de Uso")
    st.markdown(f"""
### 🚨 Como Funciona o Salvamento de Géis de Conferência?
Toda vez que você salva a matriz de um primer:
1. O sistema cria automaticamente a pasta `ssr_resultados_salvos/geis/`.
2. Salva nessa pasta uma cópia **.jpg** da imagem unida dos seus dois géis contendo a **linha do laser, as linhas das bandas e as marcações verdes**.
3. O nome do arquivo salvo inclui o nome do primer e o horário exato da gravação (ex: `ISSR_19_20250321_153000.jpg`).

### 📝 Controles Rápidos do Visualizador
| Ação | Como fazer |
|------|------------|
| **Zoom** | Roda do mouse (Scroll Wheel) |
| **Mover a Imagem** | Botão do meio do mouse **ou** Segurar barra de Espaço e arrastar |
| **Criar Nova Coluna** | Ativar o modo azul de edição e clicar com Botão Esquerdo no gel |
| **Mover Coluna** | Segurar e arrastar a linha correspondente |
| **Apagar Linha de Coluna**| Clique com Botão Direito do mouse em cima da linha |
| **Criar Banda** | Botão Esquerdo no gel (Fora das bandas existentes) |
| **Marcar Banda (0/1)** | Clicar com o Botão Esquerdo sobre o marcador verde |
| **Deletar Banda** | Botão Direito sobre a banda a ser apagada |
    """)
    st.divider()
    st.success("✅ SSR Pro v28.5 — Fotos dos géis com marcações salvas automaticamente no PC para conferências futuras.")