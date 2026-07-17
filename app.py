import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Tekstilbox & EasyExpo B2B Lead Finder",
    page_icon="🎯",
    layout="wide"
)

# 1. KULLANICI GİRİŞ KONTROLÜ (Şirket İçi Güvenlik)
# Kolaylık olması açısından şifreleri doğrudan buraya tanımladık. İstediğiniz gibi değiştirebilirsiniz.
AUTHORIZED_USERS = {
    "admin": "tekstilbox2026",
    "satis1": "satis2026",
    "satis2": "ees2026"
}

def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.subheader("🔑 Tekstilbox & EasyExpo - B2B Lead Finder Giriş Paneli")
    col1, col2 = st.columns(2)
    with col1:
        username = st.text_input("Kullanıcı Adı")
        password = st.text_input("Şifre", type="password")
        if st.button("Giriş Yap 🚀"):
            if username in AUTHORIZED_USERS and AUTHORIZED_USERS[username] == password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Hatalı kullanıcı adı veya şifre!")
    return False

if check_password():
    # Sidebar - Ayarlar ve API Anahtarı
    st.sidebar.image("https://www.tekstilbox.com/wp-content/uploads/2021/04/cropped-tekstilbox-logo-black-1.png", width=200)
    st.sidebar.title("⚙️ Ayarlar")
    
    # Kullanıcının kod bilmeden kendi API anahtarını girebilmesi için alan
    serpapi_key = st.sidebar.text_input(
        "SerpApi Anahtarınız (API Key)", 
        type="password", 
        help="serpapi.com sitesinden ücretsiz veya ücretli aldığınız anahtarı buraya yapıştırın."
    )
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📌 Nasıl Çalışır?")
    st.sidebar.info(
        "1. Ülkeyi ve hedef kitleyi seçin.\n"
        "2. Kelimelerinizi belirleyip aramayı başlatın.\n"
        "3. Sistem Google üzerinden LinkedIn sayfalarını yerel dilde tarar.\n"
        "4. Excel (.xlsx) dosyasını indirip firmaları inceleyin."
    )

    st.title("🎯 Tekstilbox & EasyExpo B2B Arama Motoru")
    st.write("Hedef kitle bazlı, akıllı yerel dil çevirili ve filtreli müşteri bulma programı.")

    # 2. SÖZLÜK VE KELİME KÜTÜPHANESİ
    # İleride yeni diller veya kelimeler eklemek için bu yapıyı genişletebilirsiniz
    LOCAL_DICTIONARY = {
        "FR": {
            "shopfitter": "agenceur",
            "lightbox": "caisson lumineux",
            "retail design": "architecture commerciale",
            "shopfitting": "agencement de magasin",
            "signage": "signalétique",
            "sign": "enseigne",
            "profile": "profilé",
            "fabric": "tissu",
            "stand": "stand d'exposition"
        },
        "DE": {
            "shopfitter": "ladenbauer",
            "lightbox": "leuchtkasten",
            "retail design": "ladenbau design",
            "shopfitting": "ladeneinrichtung",
            "signage": "beschilderung",
            "sign": "schilder",
            "profile": "profile",
            "fabric": "textilfalwand",
            "stand": "messestand"
        },
        "NL": {
            "shopfitter": "winkelinterieurbouw",
            "lightbox": "lichtbak",
            "retail design": "retail design",
            "shopfitting": "winkelindeling",
            "signage": "bewegwijzering",
            "sign": "borden",
            "profile": "profiel",
            "fabric": "stof",
            "stand": "beursstand"
        },
        "IT": {
            "shopfitter": "allestitore",
            "lightbox": "cassonetto luminoso",
            "retail design": "progettazione negozi",
            "shopfitting": "allestimento negozi",
            "signage": "segnaletica",
            "sign": "insegna",
            "profile": "profilo",
            "fabric": "tessuto",
            "stand": "stand fieristico"
        }
    }

    # 3. DİNAMİK SEÇİM ARABİRİMİ
    col_left, col_right = st.columns([1, 2])

    with col_left:
        country = st.selectbox(
            "1. Hedef Ülke Seçin", 
            ["Fransa (FR)", "Almanya (DE)", "İngiltere (GB)", "Hollanda (NL)", "İtalya (IT)"]
        )
        country_code = country.split("(")[1].replace(")", "")

        hedef_kitle = st.selectbox(
            "2. Hedef Kitle Seçin", 
            ["Shopfitter", "Signage/Sign", "Producer", "Print house", "Modular Exhibition Systems"]
        )

        keywords_to_search = []

        # Koşullu Seçenekler
        if hedef_kitle == "Shopfitter":
            keywords_to_search = st.multiselect(
                "Aranacak Anahtar Kelimeler (En az 1 adet seçin)",
                ["shopfitter", "lightbox", "retail design", "shopfitting"],
                default=["shopfitter", "lightbox"]
            )
            
        elif hedef_kitle == "Signage/Sign":
            keywords_to_search = st.multiselect(
                "Aranacak Anahtar Kelimeler",
                ["lightbox", "signage", "sign"],
                default=["lightbox", "sign"]
            )

        elif hedef_kitle == "Producer":
            producer_type = st.radio("Ürün Grubu Seçin", ["Profil", "Kumaş", "LED"])
            if producer_type == "Profil":
                keywords_to_search = st.multiselect("Profil Seçenekleri", ["double side", "single side", "lightbox", "profile"])
            elif producer_type == "Kumaş":
                keywords_to_search = st.multiselect("Kumaş Seçenekleri", ["SEG Fabric", "UV Fabric", "Dye-Sub Fabric", "Sublimation Fabric"])
            elif producer_type == "LED":
                keywords_to_search = st.multiselect("LED Seçenekleri", ["module led", "lightbox led"])

        elif hedef_kitle == "Print house":
            keywords_to_search = st.multiselect(
                "Baskı Teknolojileri Seçin",
                ["Digital Printing", "UV printing", "Dye-Sub printing", "Sublimation printing", "large format printing"]
            )

        elif hedef_kitle == "Modular Exhibition Systems":
            moduler_type = st.radio("Sistem Seçin", ["xPowall+", "EES"])
            if moduler_type == "xPowall+":
                keywords_to_search = st.multiselect("Marka/Uyum Seçenekleri", ["beMatrix", "Aluvision"])
            elif moduler_type == "EES":
                keywords_to_search = st.multiselect("Sistem Seçenekleri", ["Modular Exhibition Systems", "Modular Expo Systems", "Stand"])

    with col_right:
        st.subheader("🔍 Arama Bilgileri ve Önizleme")
        
        if not keywords_to_search:
            st.warning("Lütfen sol taraftan en az bir anahtar kelime seçin.")
        else:
            # Arama Sorgusunu Oluşturma (Yerel Dil Entegrasyonu)
            query_parts = []
            local_dict = LOCAL_DICTIONARY.get(country_code, {})
            
            for kw in keywords_to_search:
                local_kw = local_dict.get(kw.lower())
                if local_kw and local_kw != kw.lower():
                    # Hem İngilizce hem Yerel Dilde Aratır: ("lightbox" OR "caisson lumineux")
                    query_parts.append(f'("{kw}" OR "{local_kw}")')
                else:
                    query_parts.append(f'"{kw}"')
                    
            search_query = f"site:linkedin.com/company/ " + " AND ".join(query_parts)
            
            st.code(f"Oluşturulan Google Arama Sorgusu:\n{search_query}", language="text")
            
            # Arama Butonu
            if st.button("Müşterileri Bul 🚀", use_container_width=True):
                if not serpapi_key:
                    st.error("⚠️ Lütfen sol menüdeki 'Ayarlar' kısmından SerpApi anahtarınızı girin!")
                else:
                    st.info("🔄 Google verileri taranıyor, bu işlem yaklaşık 5-10 saniye sürebilir...")
                    
                    # API Parametreleri
                    url = "https://serpapi.com/search"
                    params = {
                        "engine": "google",
                        "q": search_query,
                        "gl": country_code.lower(),
                        "num": "100",  # Google'dan en iyi 100 sonucu çekiyoruz
                        "api_key": serpapi_key
                    }
                    
                    try:
                        response = requests.get(url, params=params)
                        data = response.json()
                        results = data.get("organic_results", [])
                        
                        leads = []
                        for result in results:
                            linkedin_url = result.get("link")
                            if "linkedin.com/company/" in linkedin_url:
                                raw_title = result.get("title")
                                # Başlıktan temiz şirket ismi çıkarma (Gereksiz kısımları temizler)
                                clean_title = raw_title.split(":")[0].split("|")[0].split("-")[0].strip()
                                
                                # Web sitesi çıkarma denemesi (snippet veya Google verisinden)
                                snippet = result.get("snippet", "")
                                
                                leads.append({
                                    "Ülke": country_code,
                                    "Firma İsmi": clean_title,
                                    "Web Sitesi": "Lütfen siteyi incelemek için aratın veya LinkedIn sayfasından bakın",
                                    "LinkedIn Sayfası Linki": linkedin_url
                                })
                        
                        if leads:
                            df = pd.DataFrame(leads)
                            st.success(f"🎉 Başarılı! Toplam {len(df)} adet potansiyel firma listelendi.")
                            st.dataframe(df, use_container_width=True)
                            
                            # Excel'i bellekte profesyonelce biçimlendirme
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df.to_excel(writer, index=False, sheet_name="Müşteri Listesi")
                                
                                # Excel Görsel Tasarımı (openpyxl kullanarak)
                                workbook = writer.book
                                worksheet = writer.sheets["Müşteri Listesi"]
                                
                                # Başlık satırı stili (Lacivert arka plan, beyaz kalın yazı)
                                header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                                header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
                                center_align = Alignment(horizontal="center", vertical="center")
                                left_align = Alignment(horizontal="left", vertical="center")
                                
                                # Hücre kenarlıkları
                                thin_border = Border(
                                    left=Side(style='thin', color='D3D3D3'),
                                    right=Side(style='thin', color='D3D3D3'),
                                    top=Side(style='thin', color='D3D3D3'),
                                    bottom=Side(style='thin', color='D3D3D3')
                                )
                                
                                # Başlıkları stillendir
                                for col_idx, col in enumerate(df.columns, 1):
                                    cell = worksheet.cell(row=1, column=col_idx)
                                    cell.fill = header_fill
                                    cell.font = header_font
                                    cell.alignment = center_align
                                    cell.border = thin_border
                                
                                # Verileri ve sütun genişliklerini ayarla
                                for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1, min_col=1, max_col=4):
                                    for cell in row:
                                        cell.font = Font(name="Arial", size=10)
                                        cell.border = thin_border
                                        if cell.column == 1:  # Ülke sütununu ortala
                                            cell.alignment = center_align
                                        else:
                                            cell.alignment = left_align
                                
                                # Sütun genişliklerini otomatik genişlet
                                for col in worksheet.columns:
                                    max_len = max(len(str(cell.value or '')) for cell in col)
                                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                                    
                            excel_data = output.getvalue()
                            
                            st.download_button(
                                label="📥 Excel Dosyasını İndir (.xlsx)",
                                data=excel_data,
                                file_name=f"B2B_Leads_{hedef_kitle}_{country_code}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.warning("⚠️ Seçtiğiniz kriterlere uygun şirket bulunamadı. Anahtar kelimeleri azaltmayı veya değiştirmeyi deneyebilirsiniz.")
                            
                    except Exception as e:
                        st.error(f"Sorgu sırasında bir hata oluştu: {e}")
