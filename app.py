import io
import json
import re
import time
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
from PIL import Image
from supabase import create_client
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Gemini is optional: the app still runs without it.
try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None


st.set_page_config(
    page_title="AgroSmart - Gestão de NFP-e, GTA & CAR",
    page_icon="🌾",
    layout="wide",
)

st.markdown(
    """
    <style>
    .stApp { background: linear-gradient(135deg, #f4f6f0 0%, #e8efe2 100%); }
    h1 { color: #1b4332 !important; font-weight: 700; }
    h2, h3 { color: #2d6a4f !important; }
    [data-testid="stSidebar"] { background-color: #d8f3dc; }
    [data-testid="stFileUploader"] {
        background: #fff; padding: 12px; border-radius: 10px;
        border: 1px solid #b7e4c7;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "lancamentos" not in st.session_state:
    st.session_state.lancamentos = []


TIPOS_DOCUMENTO = {
    "Nota Fiscal de Produtor (NFP-e)": "nota",
    "Guia de Trânsito Animal (GTA)": "gta",
    "Cadastro Ambiental Rural (CAR)": "car",
}


def parse_json_safely(text):
    """Extract a JSON object from a model response."""
    if not text:
        return None
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        result = json.loads(cleaned)
        return result if isinstance(result, dict) else None
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
        if match:
            try:
                result = json.loads(match.group(0))
                return result if isinstance(result, dict) else None
            except json.JSONDecodeError:
                return None
    return None


def extract_pdf_text(file_bytes):
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            return "\n".join((page.extract_text() or "") for page in pdf.pages).strip()
    except Exception:
        return ""


def extract_fields_from_text(text, doc_type):
    """Conservative local extraction. Missing values remain blank; no invented data."""
    if not text:
        return None

    def first(patterns):
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    if doc_type == "nota":
        data = {
            "emitente_nome": first([
                r"(?:Raz[aã]o Social|Nome/Raz[aã]o Social|Emitente)\s*[:\-]?\s*([^\n]+)",
            ]),
            "emitente_cpf_cnpj": first([
                r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})",
                r"(\d{3}\.\d{3}\.\d{3}-\d{2})",
            ]),
            "numero_nf": first([
                r"(?:N[ºo°]\s*(?:da\s*)?NF-e|N[úu]mero da Nota|N[ºo°]\s*Nota|NF-e)\s*[:#\-]?\s*([0-9]{1,})",
            ]),
            "valor_total_nota": first([
                r"(?:Valor Total da Nota|Valor Total|Total da Nota)\s*[:R$\s]*([0-9.]+,[0-9]{2}|[0-9]+\.[0-9]{2})",
            ]),
        }
    elif doc_type == "gta":
        data = {
            "numero_gta": first([
                r"(?:N[ºo°]\s*(?:da\s*)?GTA|N[úu]mero da GTA|GTA)\s*[:#\-]?\s*([A-Z0-9./\-]+)",
            ]),
            "emitente": first([
                r"(?:Produtor|Propriet[aá]rio|Emitente|Nome do produtor)\s*[:\-]?\s*([^\n]+)",
            ]),
            "quantidade_cabecas": first([
                r"(?:Quantidade de animais|Quantidade|N[ºo°]\s*de animais|Total de animais)\s*[:\-]?\s*([0-9]+)",
            ]),
            "destino_uf_municipio": first([
                r"(?:Munic[ií]pio de destino|Destino)\s*[:\-]?\s*([^\n]+)",
            ]),
        }
    else:
        data = {
            "numero_car": first([
                r"([A-Z]{2}-\d{5,}-[A-F0-9]+)",
                r"(?:Registro do CAR|N[ºo°]\s*do CAR|CAR)\s*[:#\-]?\s*([A-Z0-9./\-]+)",
            ]),
            "proprietario": first([
                r"(?:Nome do Im[oó]vel|Propriet[aá]rio|Possuidor|Nome do propriet[aá]rio)\s*[:\-]?\s*([^\n]+)",
            ]),
            "area_hectares": first([
                r"(?:[ÁA]rea total|[ÁA]rea do im[oó]vel|[ÁA]rea)\s*[:\-]?\s*([0-9.]+,[0-9]+|[0-9]+\.[0-9]+)\s*(?:ha|hectares)?",
            ]),
            "status_car": first([
                r"(?:Situa[cç][aã]o do CAR|Status do CAR|Situa[cç][aã]o)\s*[:\-]?\s*([^\n]+)",
            ]),
        }

    return data if any(str(value).strip() for value in data.values()) else None


def get_gemini_client(api_key):
    if not api_key or genai is None:
        return None
    try:
        return genai.Client(api_key=api_key)
    except Exception:
        return None


def extract_with_gemini(file_bytes, filename, doc_type, client):
    if client is None or types is None:
        return None, "Gemini não está configurado."
    prompts = {
        "nota": (
            'Extraia os dados desta Nota Fiscal de Produtor. Retorne apenas JSON válido com '
            'as chaves "emitente_nome", "emitente_cpf_cnpj", "numero_nf", "valor_total_nota". '
            'Se um campo não estiver legível ou não existir, use string vazia. Não invente dados.'
        ),
        "gta": (
            'Extraia os dados desta Guia de Trânsito Animal (GTA). Retorne apenas JSON válido com '
            'as chaves "numero_gta", "emitente", "quantidade_cabecas", "destino_uf_municipio". '
            'Se um campo não estiver legível ou não existir, use string vazia. Não invente dados.'
        ),
        "car": (
            'Extraia os dados deste documento do Cadastro Ambiental Rural (CAR). Retorne apenas JSON válido com '
            'as chaves "numero_car", "proprietario", "area_hectares", "status_car". '
            'Se um campo não estiver legível ou não existir, use string vazia. Não invente dados.'
        ),
    }
    suffix = filename.lower()
    mime_type = "application/pdf" if suffix.endswith(".pdf") else (
        "image/png" if suffix.endswith(".png") else "image/jpeg"
    )
    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                prompts[doc_type],
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0,
            ),
        )
        data = parse_json_safely(getattr(response, "text", ""))
        if data and any(str(value).strip() for value in data.values()):
            return data, ""
        return None, "A IA não conseguiu extrair dados legíveis deste documento."
    except Exception as exc:
        return None, f"Falha na análise por IA: {exc}"


def format_doc_summary(item):
    raw = item.get("RawData", {})
    doc_type = item.get("Tipo Documento", "")
    if "Nota" in doc_type:
        return raw.get("numero_nf", ""), f"R$ {raw.get('valor_total_nota', '')}".strip()
    if "Guia" in doc_type:
        return raw.get("numero_gta", ""), f"{raw.get('quantidade_cabecas', '')} cabeças".strip()
    return raw.get("numero_car", ""), f"{raw.get('area_hectares', '')} ha".strip()


def make_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "AgroSmart"
    headers = [
        "Arquivo", "Data/Hora", "Produtor / Fazenda", "Tipo Documento",
        "Identificador", "Detalhes", "Status", "Dados extraídos (JSON)"
    ]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="1B4332")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D3D3D3")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for item in records:
        ident, detail = format_doc_summary(item)
        ws.append([
            item.get("NomeFicheiro", ""),
            item.get("Data/Hora", ""),
            item.get("ProdutorIdentificado", ""),
            item.get("Tipo Documento", ""),
            ident or "",
            detail or "",
            item.get("StatusAuditoria", ""),
            json.dumps(item.get("RawData", {}), ensure_ascii=False),
        ])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = [28, 20, 32, 34, 22, 26, 42, 70]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def make_domain_txt(records):
    lines = []
    for item in records:
        raw = item.get("RawData", {})
        doc_type = item.get("Tipo Documento", "")
        producer = str(item.get("ProdutorIdentificado", ""))
        if "Nota" in doc_type:
            value = str(raw.get("valor_total_nota", "")).replace(".", "").replace(",", "")
            lines.append(f"DOMINIO_NF|{producer}|NF|{raw.get('numero_nf', '')}|{value}")
        elif "Guia" in doc_type:
            lines.append(
                f"DOMINIO_GTA|{producer}|GTA|{raw.get('numero_gta', '')}|{raw.get('quantidade_cabecas', '')}"
            )
        else:
            lines.append(f"DOMINIO_CAR|{producer}|CAR|{raw.get('numero_car', '')}")
    return "\n".join(lines)


# ---------------------------
# Supabase authentication + persistence
# ---------------------------
def get_settings():
    try:
        return {
            "url": st.secrets["SUPABASE_URL"],
            "anon_key": st.secrets["SUPABASE_ANON_KEY"],
            "service_key": st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", ""),
            "app_url": st.secrets.get("APP_URL", "").rstrip("/"),
            "google_key": st.secrets.get("GOOGLE_API_KEY", ""),
        }
    except Exception:
        return {"url": "", "anon_key": "", "service_key": "", "app_url": "", "google_key": ""}


def get_auth_client(settings):
    # Keep the same client in Streamlit session so the Supabase PKCE verifier
    # can survive ordinary reruns during the password recovery flow.
    if "_agrosmart_supabase" not in st.session_state:
        st.session_state["_agrosmart_supabase"] = create_client(settings["url"], settings["anon_key"])
    return st.session_state["_agrosmart_supabase"]


def get_admin_client(settings):
    if not settings.get("service_key"):
        raise RuntimeError("Configure SUPABASE_SERVICE_ROLE_KEY nos Secrets para usar as funções administrativas.")
    return create_client(settings["url"], settings["service_key"])


def get_db_client(settings, access_token):
    client = create_client(settings["url"], settings["anon_key"])
    client.postgrest.auth(access_token)
    return client


def response_user(response):
    user = getattr(response, "user", None)
    session = getattr(response, "session", None)
    if not user or not session:
        return False
    st.session_state["ag_user_id"] = str(user.id)
    st.session_state["ag_email"] = str(getattr(user, "email", "") or "")
    st.session_state["ag_access_token"] = session.access_token
    st.session_state["ag_refresh_token"] = session.refresh_token
    app_meta = getattr(user, "app_metadata", {}) or {}
    st.session_state["ag_role"] = app_meta.get("role", "client")
    return True


def clear_login():
    for key in ("ag_user_id", "ag_email", "ag_access_token", "ag_refresh_token", "ag_role", "ag_profile"):
        st.session_state.pop(key, None)
    st.session_state.pop("lancamentos", None)


def load_profile(settings):
    token = st.session_state.get("ag_access_token", "")
    if not token:
        return None
    db = get_db_client(settings, token)
    result = db.table("profiles").select("id,email,full_name,active").eq("id", st.session_state["ag_user_id"]).limit(1).execute()
    profile = result.data[0] if result.data else None
    st.session_state["ag_profile"] = profile
    return profile


def fetch_documents(settings):
    db = get_db_client(settings, st.session_state["ag_access_token"])
    result = db.table("documents").select("id,user_id,file_name,created_at,doc_type,producer,raw_data,status,observation").order("created_at", desc=True).execute()
    records = []
    for row in result.data or []:
        try:
            date_text = datetime.fromisoformat(row.get("created_at", "").replace("Z", "+00:00")).astimezone().strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            date_text = row.get("created_at", "")
        records.append({
            "id": row.get("id"),
            "NomeFicheiro": row.get("file_name", ""),
            "Data/Hora": date_text,
            "Tipo Documento": row.get("doc_type", ""),
            "ProdutorIdentificado": row.get("producer", ""),
            "RawData": row.get("raw_data") or {},
            "StatusAuditoria": row.get("status", ""),
            "Observacao": row.get("observation", ""),
        })
    return records


def save_document(settings, item):
    db = get_db_client(settings, st.session_state["ag_access_token"])
    payload = {
        "user_id": st.session_state["ag_user_id"],
        "file_name": item.get("NomeFicheiro", ""),
        "doc_type": item.get("Tipo Documento", ""),
        "producer": item.get("ProdutorIdentificado", ""),
        "raw_data": item.get("RawData", {}),
        "status": item.get("StatusAuditoria", ""),
        "observation": item.get("Observacao", ""),
    }
    db.table("documents").insert(payload).execute()


def delete_document(settings, doc_id):
    db = get_db_client(settings, st.session_state["ag_access_token"])
    db.table("documents").delete().eq("id", doc_id).execute()



def show_login(settings, supabase):
    st.title("🌾 AgroSmart")
    st.subheader("Acesse sua conta")
    login_tab, recover_tab = st.tabs(["Entrar", "Esqueci minha senha"])

    with login_tab:
        with st.form("login_form"):
            email = st.text_input("E-mail", key="login_email").strip().lower()
            password = st.text_input("Senha", type="password")
            submitted = st.form_submit_button(
                "Entrar", type="primary", use_container_width=True
            )

        if submitted:
            try:
                response = supabase.auth.sign_in_with_password(
                    {"email": email, "password": password}
                )
                if response_user(response):
                    st.rerun()
                else:
                    st.error(
                        "Não foi possível iniciar a sessão. Confirme seu e-mail "
                        "e senha e verifique a confirmação do e-mail, se ativada."
                    )
            except Exception:
                st.error(
                    "Não foi possível entrar. Verifique o e-mail e a senha "
                    "ou tente novamente."
                )

    with recover_tab:
        st.write(
            "Informe o e-mail cadastrado. Se a conta existir, "
            "você receberá um link para criar outra senha."
        )

        with st.form("recover_form"):
            recover_email = st.text_input(
                "E-mail da conta", key="recover_email"
            ).strip().lower()
            recover_submit = st.form_submit_button(
                "Enviar link de recuperação",
                use_container_width=True,
            )

        if recover_submit:
            if not settings.get("app_url"):
                st.error(
                    "O administrador precisa configurar APP_URL "
                    "nos Secrets do Streamlit Cloud."
                )
            elif not recover_email:
                st.warning("Informe o e-mail da conta.")
            else:
                try:
                    supabase.auth.reset_password_for_email(
                        recover_email,
                        {
                            "redirect_to": (
                                f"{settings['app_url']}/?mode=recovery"
                            )
                        },
                    )
                    st.success(
                        "Se o e-mail estiver cadastrado, o link de "
                        "recuperação será enviado."
                    )
                except Exception as exc:
                    st.error(
                        f"Erro ao solicitar recuperação: {exc}"
                    )


def admin_panel(settings):
    st.subheader("🛡️ Administração de clientes")
    if not settings.get("service_key"):
        st.warning("A função administrativa exige SUPABASE_SERVICE_ROLE_KEY nos Secrets. Nunca compartilhe essa chave nem a coloque no GitHub.")
        return
    admin = get_admin_client(settings)
    with st.expander("➕ Cadastrar novo cliente", expanded=False):
        with st.form("create_client_form"):
            full_name = st.text_input("Nome do cliente")
            email = st.text_input("E-mail do cliente").strip().lower()
            password = st.text_input("Senha inicial (mínimo de 10 caracteres)", type="password")
            create = st.form_submit_button("Criar conta", type="primary")
        if create:
            if not full_name.strip() or not email or len(password) < 10:
                st.error("Preencha o nome, o e-mail e uma senha com pelo menos 10 caracteres.")
            else:
                try:
                    created = admin.auth.admin.create_user({
                        "email": email,
                        "password": password,
                        "email_confirm": True,
                        "user_metadata": {"full_name": full_name.strip()},
                    })
                    user = getattr(created, "user", None)
                    if user is None:
                        raise RuntimeError("O Supabase não retornou o usuário criado.")
                    admin.table("profiles").upsert({
                        "id": str(user.id), "email": email,
                        "full_name": full_name.strip(), "active": True,
                    }).execute()
                    st.success(f"Conta criada para {email}. Entregue a senha inicial por um canal seguro e peça ao cliente para alterá-la.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não foi possível criar a conta. Verifique se o e-mail já existe e se a chave administrativa está correta. Detalhe: {exc}")

    st.markdown("**Clientes cadastrados**")
    try:
        result = admin.table("profiles").select("id,email,full_name,active,created_at").order("created_at", desc=True).execute()
        profiles = result.data or []
    except Exception as exc:
        st.error(f"Não foi possível carregar a lista de clientes: {exc}")
        return
    if not profiles:
        st.info("Ainda não há clientes cadastrados.")
        return
    for profile in profiles:
        if profile["id"] == st.session_state["ag_user_id"]:
            continue
        cols = st.columns([3, 2, 1.3])
        cols[0].write(f"**{profile.get('full_name') or 'Sem nome'}**\n\n{profile.get('email', '')}")
        cols[1].write("Ativo" if profile.get("active", True) else "Bloqueado")
        if profile.get("active", True):
            if cols[2].button("Bloquear", key=f"block_{profile['id']}"):
                try:
                    admin.table("profiles").update({"active": False}).eq("id", profile["id"]).execute()
                    # Ban at Auth level as a second layer; RLS also checks profile active.
                    try:
                        admin.auth.admin.update_user_by_id(profile["id"], {"ban_duration": "876000h"})
                    except Exception:
                        pass
                    st.success("Acesso bloqueado.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não foi possível bloquear a conta: {exc}")
        else:
            if cols[2].button("Reativar", key=f"unblock_{profile['id']}"):
                try:
                    admin.table("profiles").update({"active": True}).eq("id", profile["id"]).execute()
                    try:
                        admin.auth.admin.update_user_by_id(profile["id"], {"ban_duration": "none"})
                    except Exception:
                        pass
                    st.success("Acesso reativado.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Não foi possível reativar a conta: {exc}")


settings = get_settings()
if not settings.get("url") or not settings.get("anon_key"):
    st.title("🌾 AgroSmart")
    st.error("Configure SUPABASE_URL e SUPABASE_ANON_KEY em Settings → Secrets no Streamlit Cloud.")
    st.code('SUPABASE_URL = "https://SEU-PROJETO.supabase.co"\nSUPABASE_ANON_KEY = "SUA_CHAVE_ANON_PUBLIC"\nSUPABASE_SERVICE_ROLE_KEY = "SUA_CHAVE_SERVICE_ROLE"\nAPP_URL = "https://agrosmartfuturodoagro.streamlit.app"\nGOOGLE_API_KEY = "SUA_CHAVE_GEMINI_OPCIONAL"', language="toml")
    st.stop()

try:
    supabase = get_auth_client(settings)
except Exception:
    st.title("🌾 AgroSmart")
    st.error("Não foi possível conectar ao Supabase. Confira SUPABASE_URL e SUPABASE_ANON_KEY nos Secrets.")
    st.stop()

# Password-recovery callback. Prefer token_hash from the customized Supabase
# recovery email template; also accept the standard PKCE ?code= callback.
if st.query_params.get("mode") == "recovery" and st.query_params.get("token_hash"):
    try:
        callback_response = supabase.auth.verify_otp({
            "token_hash": st.query_params.get("token_hash"),
            "type": st.query_params.get("type", "recovery"),
        })
        response_user(callback_response)
        st.session_state["ag_recovery_mode"] = True
    except Exception:
        st.session_state["ag_recovery_error"] = True
elif st.query_params.get("mode") == "recovery" and st.query_params.get("code"):
    try:
        callback_code = st.query_params.get("code")
        callback_response = supabase.auth.exchange_code_for_session({"auth_code": callback_code})
        response_user(callback_response)
        st.session_state["ag_recovery_mode"] = True
    except Exception:
        st.session_state["ag_recovery_error"] = True

if st.session_state.get("ag_recovery_error"):
    st.title("🌾 Recuperar senha")
    st.error("O link de recuperação expirou ou não pôde ser validado. Solicite um novo link na tela de login.")
    if st.button("Voltar ao login"):
        st.session_state.pop("ag_recovery_error", None)
        st.query_params.clear()
        st.rerun()
    st.stop()

if st.session_state.get("ag_recovery_mode"):
    st.title("🔐 Criar nova senha")
    st.write("Escolha uma senha nova para sua conta.")
    with st.form("new_password_form"):
        new_password = st.text_input("Nova senha (mínimo de 10 caracteres)", type="password")
        confirm_password = st.text_input("Confirme a nova senha", type="password")
        change_password = st.form_submit_button("Salvar nova senha", type="primary")
    if change_password:
        if len(new_password) < 10:
            st.error("A senha precisa ter pelo menos 10 caracteres.")
        elif new_password != confirm_password:
            st.error("As senhas não coincidem.")
        else:
            try:
                supabase.auth.update_user({"password": new_password})
                st.session_state.pop("ag_recovery_mode", None)
                clear_login()
                st.query_params.clear()
                st.success("Senha atualizada. Entre com sua nova senha.")
                st.rerun()
            except Exception:
                st.error("Não foi possível atualizar a senha. Solicite um novo link de recuperação e tente novamente.")
    st.stop()

if not st.session_state.get("ag_access_token"):
    show_login(settings, supabase)
    st.stop()

# Refresh token/session if needed; session is kept in Streamlit session state.
try:
    if st.session_state.get("ag_refresh_token"):
        refreshed = supabase.auth.set_session(
            st.session_state["ag_access_token"], st.session_state["ag_refresh_token"]
        )
        session = getattr(refreshed, "session", None)
        if session:
            st.session_state["ag_access_token"] = session.access_token
            st.session_state["ag_refresh_token"] = session.refresh_token
except Exception:
    clear_login()
    st.warning("Sua sessão expirou. Entre novamente.")
    st.rerun()

try:
    profile = load_profile(settings)
except Exception:
    profile = None
if not profile:
    clear_login()
    st.error("Seu perfil não foi encontrado. Peça ao administrador para verificar seu cadastro.")
    st.stop()
if not profile.get("active", True):
    clear_login()
    st.error("Esta conta está bloqueada. Entre em contato com o administrador do AgroSmart.")
    st.stop()

is_admin = st.session_state.get("ag_role") == "admin"

st.markdown(
    """
    <style>
    .stApp { background: linear-gradient(135deg, #f4f6f0 0%, #e8efe2 100%); }
    h1 { color: #1b4332 !important; font-weight: 700; }
    h2, h3 { color: #2d6a4f !important; }
    [data-testid="stSidebar"] { background-color: #d8f3dc; }
    [data-testid="stFileUploader"] { background: #fff; padding: 12px; border-radius: 10px; border: 1px solid #b7e4c7; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🌾 AgroSmart - Gestão de NFP-e, GTA & CAR")
st.caption(f"Conta: {profile.get('full_name') or st.session_state.get('ag_email')} · {'Administrador' if is_admin else 'Cliente'}")
with st.sidebar:
    st.header("👤 Minha conta")
    st.write(st.session_state.get("ag_email", ""))
    if st.button("Sair", use_container_width=True):
        try:
            supabase.auth.sign_out()
        except Exception:
            pass
        clear_login()
        st.rerun()
    st.divider()
    st.header("⚙️ Configurações & Filtros")
    doc_label = st.selectbox("Categoria do documento:", list(TIPOS_DOCUMENTO.keys()))
    st.markdown("---")
    st.subheader("🤖 Extração com IA (opcional)")
    st.caption("Sem chave Gemini, o sistema tenta extrair texto de PDFs nativos. Para imagens ou PDFs digitalizados, configure a chave no Secrets.")
    api_key = settings.get("google_key", "")
    st.markdown("---")
    status_filter = st.selectbox("Filtrar documentos:", ["Todos os registos", "Extração concluída", "Revisão necessária"])
    if st.button("🔄 Atualizar lista", use_container_width=True):
        st.rerun()

if is_admin:
    with st.expander("🛡️ Painel do administrador", expanded=False):
        admin_panel(settings)

try:
    records = fetch_documents(settings)
except Exception as exc:
    st.error("Não foi possível carregar os registros do banco. Confirme se executou o arquivo supabase_schema.sql no SQL Editor do Supabase.")
    st.caption(str(exc))
    st.stop()

st.divider()
doc_type = TIPOS_DOCUMENTO[doc_label]
st.subheader(f"📁 Carregar documentos: {doc_label}")
uploaded_files = st.file_uploader("Selecione um ou mais documentos (PDF, PNG ou JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files and st.button("🚀 Processar documentos", type="primary"):
    client = get_gemini_client(api_key)
    progress = st.progress(0)
    status_box = st.empty()
    added = 0
    for index, uploaded in enumerate(uploaded_files):
        status_box.write(f"Processando {index + 1}/{len(uploaded_files)}: {uploaded.name}")
        file_bytes = uploaded.getvalue()
        data = None
        method = ""
        if uploaded.name.lower().endswith(".pdf"):
            pdf_text = extract_pdf_text(file_bytes)
            if pdf_text:
                data = extract_fields_from_text(pdf_text, doc_type)
                if data:
                    method = "Extraído do texto do PDF — confira os campos."
        if not data and client:
            data, error = extract_with_gemini(file_bytes, uploaded.name, doc_type, client)
            if data:
                method = "Extraído com Gemini — confira os campos."
            else:
                method = error or "Não foi possível extrair os dados."
        if not data:
            data = {}
            method = method or "Nenhum dado extraído. Confira o tipo escolhido e, se necessário, configure o Gemini."
        producer = data.get("emitente_nome") or data.get("emitente") or data.get("proprietario") or ""
        status = "⚠️ Revisão necessária" if not any(str(v).strip() for v in data.values()) else "✅ Extração concluída"
        item = {
            "NomeFicheiro": uploaded.name,
            "Data/Hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Tipo Documento": doc_label,
            "ProdutorIdentificado": producer,
            "RawData": data,
            "StatusAuditoria": status,
            "Observacao": method,
        }
        try:
            save_document(settings, item)
            added += 1
        except Exception as exc:
            st.error(f"Não foi possível salvar {uploaded.name}. Nenhum registro foi confirmado no banco. Detalhe: {exc}")
        progress.progress((index + 1) / len(uploaded_files))
    status_box.empty()
    progress.empty()
    if added:
        st.success(f"Processamento finalizado: {added} registro(s) salvo(s) na sua conta.")
    st.rerun()

# Reload after processing; database RLS enforces tenant separation.
try:
    records = fetch_documents(settings)
except Exception as exc:
    st.error(f"Falha ao carregar os registros: {exc}")
    st.stop()

if status_filter == "Extração concluída":
    visible = [r for r in records if "Extração concluída" in r.get("StatusAuditoria", "")]
elif status_filter == "Revisão necessária":
    visible = [r for r in records if "Revisão necessária" in r.get("StatusAuditoria", "")]
else:
    visible = records

st.divider()
st.subheader("📊 Painel consolidado")
if not visible:
    st.info("Nenhum registro nesta conta para mostrar. Carregue um documento para começar.")
else:
    col1, col2, col3 = st.columns(3)
    col1.metric("Documentos exibidos", len(visible))
    col2.metric("Extrações concluídas", sum("Extração concluída" in r.get("StatusAuditoria", "") for r in visible))
    col3.metric("Precisam de revisão", sum("Revisão necessária" in r.get("StatusAuditoria", "") for r in visible))
    table_rows = []
    for item in visible:
        ident, detail = format_doc_summary(item)
        table_rows.append({
            "Arquivo": item.get("NomeFicheiro", ""),
            "Produtor / Fazenda": item.get("ProdutorIdentificado", "") or "Não identificado",
            "Tipo": item.get("Tipo Documento", ""),
            "Identificador": ident or "Não identificado",
            "Detalhes": detail or "Não identificado",
            "Status": item.get("StatusAuditoria", ""),
        })
    st.dataframe(pd.DataFrame(table_rows), use_container_width=True, hide_index=True)
    
    st.subheader("🗑️ Remover documento")

    selected_doc = st.selectbox(
        "Selecione o documento que deseja excluir:",
        options=visible,
        format_func=lambda item: (
            f"{item.get('NomeFicheiro', 'Sem nome')} — "
            f"{item.get('Data/Hora', '')} — "
            f"{item.get('Tipo Documento', '')}"
        ),
        key="doc_to_delete",
    )

    confirm_delete = st.checkbox(
        f"Confirmo a exclusão do documento "
        f"{selected_doc.get('NomeFicheiro', '')}",
        key=f"confirm_delete_{selected_doc['id']}",
    )

    if st.button(
        "🗑️ Excluir documento selecionado",
        disabled=not confirm_delete,
        key=f"delete_{selected_doc['id']}",
    ):
        try:
            delete_document(settings, selected_doc["id"])
            st.success("Documento excluído com sucesso!")
            st.rerun()
        except Exception as exc:
            st.error(f"Não foi possível excluir o documento: {exc}")

    with st.expander("Ver dados extraídos e observações"):
        for item in visible:
            st.markdown(f"**{item.get('NomeFicheiro', '')}**")
            st.caption(item.get("Observacao", ""))
            st.json(item.get("RawData", {}))
    col_a, col_b = st.columns(2)
    with col_a:
        st.download_button("📥 Baixar relatório Excel", data=make_excel(visible), file_name=f"Relatorio_AgroSmart_{datetime.now():%Y%m%d_%H%M}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
    with col_b:
        st.download_button("📄 Baixar layout Domínio (.txt)", data=make_domain_txt(visible), file_name=f"Importacao_Dominio_{datetime.now():%Y%m%d_%H%M}.txt", mime="text/plain", use_container_width=True)

st.caption("Importante: a extração automática pode errar. Confira os dados nos documentos originais antes de usar para decisões fiscais, sanitárias ou ambientais.")
