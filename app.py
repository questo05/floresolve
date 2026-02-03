import streamlit as st
import pandas as pd
import io
import xlsxwriter
from solver import run_solver

# --- PAGINA CONFIGURATIE ---
st.set_page_config(page_title="Orkest Planner", page_icon="🎻", layout="wide")

# Verberg de link-icoontjes naast headers voor een strakker uiterlijk
st.markdown("""
    <style>
    [data-testid="stHeaderAction"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

# --- DE TABS ---
tab_setup, tab_planner = st.tabs(["🛠️ Admin Setup", "🎻 Planner"])

# ==========================================
# TAB 1: ADMIN SETUP
# ==========================================
with tab_setup:
    st.header("Start een nieuw orkest-project")
    st.write("Volg deze stappen om jouw eigen omgeving op te zetten:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.info("Klik op de knop hieronder. Google vraagt je om een kopie te maken. Dit wordt JOUW formulier.")
        st.subheader("Stap 1: Maak jouw Formulier")
        
        # Link om een kopie te maken van het Google Form
        master_form_link = "https://docs.google.com/forms/d/1cc3QyR7NvpHoLnNRuR-d-SUh3XzIhsFAwpBnFq-hYuM/copy"
        
        st.link_button("📝 Maak mijn Google Formulier", master_form_link, type="primary")
        
        st.markdown("""
        **Wat moet ik doen?**
        1. Klik de knop en kies **'Kopie maken'**.
        2. Pas de titel aan (bv. "Zomerconcert 2026").
        3. Klik in je nieuwe formulier rechtsboven op **'Verzenden'** 🔗.
        4. Stuur die link naar je muzikanten.
        """)

    with col2:
        st.write("") 
        st.write("") 
        st.write("")
        st.write("") 
        st.write("")
        st.subheader("Stap 2: Verzamel Antwoorden")
        st.write("") 
        st.write("")
        st.write("Hebben je muzikanten ingevuld?")
        st.markdown("""
        1. Ga naar je formulier.
        2. Klik op **Antwoorden**.
        3. Klik op het groene icoontje (Spreadsheets).
        4. In Sheets: **Bestand** -> **Downloaden** -> **Excel (.xlsx)**.
        5. Ga naar het tabblad **'Planner'** hierboven en upload dat bestand!
        """)
    
    st.success("Ga naar het tabblad 'Planner' om verder te gaan.")

# ==========================================
# TAB 2: DE PLANNER
# ==========================================
with tab_planner:
    st.title("🎻 Orkest Planner van Flore")
    st.markdown("Upload één Excel met namen, instrumenten, wensen en scores (3=Graag, 2=Kan, 1=Liever niet, 0=Nee).")

    # --- A. TEMPLATE BUILDER ---
    with st.expander("🛠️ Genereer hier een leeg excel sjabloon (voor handmatige invoer)"):
        st.write("Vul hier in hoe je orkest eruit ziet.")
        
        num_shows = st.number_input("Aantal Shows/Projecten", min_value=1, value=1, step=1)
        
        # Standaard tabelletje
        default_setup = pd.DataFrame([
            {"Instrument": "Viool 1", "Aantal Muzikanten": 10},
            {"Instrument": "Viool 2", "Aantal Muzikanten": 8},
            {"Instrument": "Altviool", "Aantal Muzikanten": 6},
            {"Instrument": "Cello", "Aantal Muzikanten": 6},
            {"Instrument": "Contrabas", "Aantal Muzikanten": 4},
        ])
        
        config_df = st.data_editor(default_setup, num_rows="dynamic")
        
        if st.button("Genereer mijn Excel-sjabloon"):
            rows = []
            show_cols = [f"Show {i+1}" for i in range(num_shows)]
            
            for index, row in config_df.iterrows():
                instr = row['Instrument']
                count = int(row['Aantal Muzikanten'])
                
                for i in range(count):
                    new_row = {"Naam": f"Naam {i+1}", "Instrument": instr, "Wens": "-"}
                    for show in show_cols:
                        new_row[show] = "-"
                    rows.append(new_row)
            
            df_gen = pd.DataFrame(rows)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_gen.to_excel(writer, index=False, sheet_name='Rooster')
            buffer.seek(0)
            
            st.success(f"Gedaan! Een sjabloon met {len(df_gen)} muzikanten staat klaar.")
            st.download_button("📥 Download Sjabloon", buffer, "Orkest_Data.xlsx")

    # --- B. FILE UPLOAD ---
    uploaded_file = st.file_uploader("Upload je Rooster Excel", type=['xlsx'])

    # !!! ALLES HIERONDER IS INGESPRONGEN OMDAT HET 'df' NODIG HEEFT !!!
    if uploaded_file:
        df = pd.read_excel(uploaded_file)

        # --- DATA SCHOONMAAK ---
        df = df.dropna(subset=['Naam', 'Instrument'])
        cols_to_drop = [c for c in df.columns if 'Tijdstempel' in c or 'Timestamp' in c]
        df = df.drop(columns=cols_to_drop, errors='ignore')
        
        # Splits instrumenten (voor multi-instrumentalisten)
        df['Instrument'] = df['Instrument'].astype(str).str.split(', ')
        df = df.explode('Instrument')
        
        required = ['Naam', 'Instrument', 'Wens']
        if not all(col in df.columns for col in required):
            st.error(f"Je Excel mist verplichte kolommen: {required}")
            st.stop()

        st.success("Bestand ingelezen! ✅")
        
        # --- C. LIMIETEN INSTELLEN (ZIJBALK) ---
        st.sidebar.header("🎛️ Bezetting")
        st.sidebar.write("Hoeveel mensen mogen er max per show spelen?")

        instrumenten = df['Instrument'].unique()
        limits = {} 
        
        for instr in instrumenten:
            aantal_beschikbaar = len(df[df['Instrument'] == instr])
            limits[instr] = st.sidebar.number_input(
                f"Max {instr}", 
                min_value=0, 
                max_value=aantal_beschikbaar, 
                value=0,
                key=f"limit_{instr}"
            )

        # --- D. EXTRA REGELS ---
        st.divider()
        col1, col2 = st.columns([1, 2])
        
        niet_shows = ['Naam', 'Instrument', 'Wens', 'Label', 'Resource_ID', 'Totaal']
        shows = [c for c in df.columns if c not in niet_shows and c not in ['Tijdstempel', 'Timestamp']]

        with col1:
            st.subheader("Extra Regels")
            rule_type = st.selectbox("Type regel", 
                                     ["-", 
                                      "Niet Samen (Conflict)", 
                                      "Altijd Samen", 
                                      "Persoon doet ALLE shows",
                                      "Persoon doet specifieke show"])

            if 'regels' not in st.session_state:
                st.session_state['regels'] = []
            
            regels = st.session_state['regels']
            
            # --- Invoer velden afhankelijk van regel type ---
            if rule_type == "Persoon doet ALLE shows":
                p1 = st.selectbox("Wie moet alles spelen?", df['Naam'].unique())
                if st.button("Voeg Regel Toe"):
                    regels.append({'type': 'must_all', 'p1': p1})
                    st.session_state['regels'] = regels
                    st.success(f"Regel: {p1} speelt alles.")

            elif rule_type == "Persoon doet specifieke show":
                p1 = st.selectbox("Wie?", df['Naam'].unique())
                target_show = st.selectbox("Welke show?", shows)
                if st.button("Voeg Regel Toe"):
                    regels.append({'type': 'force_show', 'p1': p1, 'show': target_show})
                    st.session_state['regels'] = regels
                    st.success(f"Regel: {p1} doet {target_show}.")
            
            elif rule_type == "Minimaal aantal shows (Per Persoon)":
                 p1 = st.selectbox("Wie?", df['Naam'].unique())
                 max_shows_totaal = len(shows) 
                 min_count = st.number_input("Minimum aantal shows", min_value=1, max_value=max_shows_totaal, value=1)
                 if st.button("Voeg Regel Toe"):
                    regels.append({'type': 'min_shows', 'p1': p1, 'count': min_count})
                    st.session_state['regels'] = regels
                    st.success(f"Regel: {p1} doet minimaal {min_count} shows.")

            elif rule_type != "-": 
                p1 = st.selectbox("Persoon 1", df['Naam'].unique())
                p2 = st.selectbox("Persoon 2", df['Naam'].unique())
                
                if p1 != p2:
                    if st.button("Voeg Regel Toe"):
                        type_code = 'conflict' if "Niet" in rule_type else 'samen'
                        regels.append({'type': type_code, 'p1': p1, 'p2': p2})
                        st.session_state['regels'] = regels
                        st.success("Regel toegevoegd!")
                else:
                    if p1 == p2 and rule_type != "-":
                         st.warning("Kies twee verschillende personen.")

        with col2:
            if regels:
                st.write("**Actieve Regels:**")
                for i, r in enumerate(regels):
                    c_txt, c_btn = st.columns([4, 1])
                    with c_txt:
                        if r['type'] == 'must_all': st.caption(f"🔒 **{r['p1']}** speelt alles")
                        elif r['type'] == 'force_show': st.caption(f"📍 **{r['p1']}** doet {r['show']}")
                        elif r['type'] == 'min_shows': st.caption(f"📉 **{r['p1']}** min. {r['count']} shows")
                        elif r['type'] == 'conflict': st.caption(f"⚡ **{r['p1']}** & **{r['p2']}** NIET samen")
                        elif r['type'] == 'samen': st.caption(f"🔗 **{r['p1']}** & **{r['p2']}** ALTIJD samen")
                    with c_btn:
                        if st.button("🗑️", key=f"del_{i}"):
                            regels.pop(i)
                            st.session_state['regels'] = regels
                            st.rerun()
                
                st.divider()
                if st.button("Alles Wissen", type="secondary"):
                    st.session_state['regels'] = []
                    st.rerun()

        # --- E. DE GROTE KNOP & RESULTAAT ---
        st.divider()
        
        # 1. Rekenwerk
        if st.button("🚀 Genereer Planning", type="primary"):
            with st.spinner("Puzzelen..."):
                status, result = run_solver(df, limits, regels)
                
                # Opslaan in geheugen
                st.session_state['oplossing_status'] = status
                st.session_state['oplossing_df'] = result
                
                # Oude handmatige bewerkingen wissen bij nieuwe berekening
                if 'bewerkte_df' in st.session_state:
                    del st.session_state['bewerkte_df']
                
                if status == "Optimal":
                    st.balloons()

        # 2. Weergave
        if 'oplossing_df' in st.session_state:
            
            status = st.session_state['oplossing_status']
            
            # --- HULPFUNCTIE VOOR EXCEL MAKEN ---
            def maak_excel(dataframe, is_checkbox_data=False):
                buffer = io.BytesIO()
                export_df = dataframe.copy()
                
                if is_checkbox_data:
                    bool_cols = [c for c in export_df.columns if export_df[c].dtype == bool]
                    for col in bool_cols:
                        export_df[col] = export_df[col].apply(lambda x: '✅' if x else '.')

                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Rooster')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Rooster']
                    format_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
                    format_center = workbook.add_format({'align': 'center'})
                    
                    worksheet.set_column('A:A', 20) 
                    worksheet.set_column('B:B', 15)
                    worksheet.set_column('C:Z', 12, format_center)
                    
                    worksheet.conditional_format('C2:Z100', {
                        'type': 'cell', 'criteria': 'equal to', 'value': '"✅"', 'format': format_green
                    })
                buffer.seek(0)
                return buffer

            # --- PREPARATIE EDITOR ---
            if 'tabel_versie' not in st.session_state:
                st.session_state['tabel_versie'] = 0

            if 'bewerkte_df' not in st.session_state:
                base_df = st.session_state['oplossing_df'].copy()
                
                volgorde_map = {naam: i for i, naam in enumerate(instrumenten)}
                base_df['_sort_index'] = base_df['Instrument'].map(volgorde_map)
                base_df = base_df.sort_values(by=['_sort_index', 'Naam'])
                base_df = base_df.drop(columns=['_sort_index'])

                for col in shows:
                    base_df[col] = base_df[col].apply(lambda x: True if x == '✅' else False)
                
                st.session_state['bewerkte_df'] = base_df

            df_to_show = st.session_state['bewerkte_df']

            if status == "Optimal":
                
                # HEADER MET KNOPPEN
                c1, c2, c3 = st.columns([2, 1, 1])
                with c1: st.subheader("Het Resultaat")
                with c2: 
                    st.write("") 
                    buffer_orig = maak_excel(st.session_state['oplossing_df'], is_checkbox_data=False)
                    st.download_button("📥 Origineel", buffer_orig, "Rooster_Origineel.xlsx", use_container_width=True)
                with c3:
                    st.write("") 
                    def reset_alles():
                        if 'bewerkte_df' in st.session_state: del st.session_state['bewerkte_df']
                        st.session_state['tabel_versie'] += 1
                    st.button("🔄 Reset Wijzigingen", type="secondary", use_container_width=True, on_click=reset_alles)

                st.info("💡 **Batch Mode:** Je kunt hieronder het rooster aanpassen. Klik op **'Opslaan'** om te controleren.")

                column_config = {
                    "Naam": st.column_config.TextColumn(disabled=True),
                    "Instrument": st.column_config.TextColumn(disabled=True),
                    "Totaal": st.column_config.NumberColumn(disabled=True)
                }
                for s in shows:
                    column_config[s] = st.column_config.CheckboxColumn(s, default=False)

                # FORMULIER
                with st.form("rooster_form"):
                    current_key = f"editor_{st.session_state['tabel_versie']}"
                    edited_df = st.data_editor(
                        df_to_show, use_container_width=True, height=600,
                        key=current_key, column_config=column_config, hide_index=True
                    )
                    st.write("")
                    submit_btn = st.form_submit_button("💾 Wijzigingen Controleren & Opslaan", type="primary")

                # LOGICA NA OPSLAAN
                if submit_btn:
                    st.session_state['bewerkte_df'] = edited_df
                    fouten_log = [] 
                    
                    # --- CHECK 1: LIMIETEN (Capaciteit) ---
                    for show in shows:
                        for instr, max_aantal in limits.items():
                            if max_aantal > 0:
                                df_instr = edited_df[edited_df['Instrument'] == instr]
                                count = df_instr[show].sum()
                                if count > max_aantal:
                                    fouten_log.append(f"⚠️ **Capaciteit:** Bij {show} zitten er **{count}** {instr}s (Max = {max_aantal}).")

                    # --- CHECK 2: FYSIEKE ONMOGELIJKHEID (De Octopus Check 🐙) ---
                    # We kijken per persoon of hij/zij meer dan 1 vinkje heeft per show
                    unieke_namen = edited_df['Naam'].unique()
                    for naam in unieke_namen:
                        persoon_rows = edited_df[edited_df['Naam'] == naam]
                        for s in shows:
                            # Tel hoeveel vinkjes deze persoon in deze show heeft staan
                            aantal_instrumenten = persoon_rows[s].sum()
                            
                            if aantal_instrumenten > 1:
                                # We zoeken even op welke instrumenten het zijn voor de duidelijkheid
                                welke_instr = persoon_rows[persoon_rows[s]]['Instrument'].tolist()
                                welke_str = " en ".join(welke_instr)
                                fouten_log.append(f"🐙 **Onmogelijk:** {naam} speelt in {s} op **{aantal_instrumenten}** instrumenten tegelijk ({welke_str}).")

                    # --- CHECK 3: REGELS (User Defined) ---
                    for r in regels:
                        try:
                            p1_rows = edited_df[edited_df['Naam'] == r.get('p1')]
                            
                            if r['type'] == 'conflict':
                                p2_rows = edited_df[edited_df['Naam'] == r.get('p2')]
                                # Heeft P1 ergens een vinkje? EN P2 ook?
                                p1_active = p1_rows[shows].any(axis=0)
                                p2_active = p2_rows[shows].any(axis=0)
                                if (p1_active & p2_active).any():
                                    fouten_log.append(f"⚡ Conflict: {r['p1']} & {r['p2']} samen ingedeeld.")

                            elif r['type'] == 'samen':
                                p2_rows = edited_df[edited_df['Naam'] == r.get('p2')]
                                p1_active = p1_rows[shows].any(axis=0)
                                p2_active = p2_rows[shows].any(axis=0)
                                if not p1_active.equals(p2_active):
                                    fouten_log.append(f"🔗 Samen: {r['p1']} en {r['p2']} lopen niet gelijk.")

                            elif r['type'] == 'must_all':
                                p1_active_per_show = p1_rows[shows].any(axis=0)
                                if not p1_active_per_show.all():
                                    missing_shows = p1_active_per_show[~p1_active_per_show].index.tolist()
                                    fouten_log.append(f"🔒 **Verplicht:** {r['p1']} mist: {', '.join(missing_shows)}")

                            elif r['type'] == 'force_show':
                                target = r['show']
                                if not p1_rows[target].any():
                                     fouten_log.append(f"📍 **Verplicht:** {r['p1']} mist in {target}.")

                            elif r['type'] == 'min_shows':
                                count = p1_rows[shows].sum().sum()
                                if count < r['count']:
                                    fouten_log.append(f"📉 **Minimum:** {r['p1']} heeft {count} shows (min {r['count']}).")
                        except Exception as e:
                            pass

                    # RESULTAAT TONEN
                    if fouten_log:
                        st.error("🛑 **Let op! Regels overtreden:**")
                        for f in fouten_log: st.write(f)
                    else:
                        st.success("✅ Alles opgeslagen! Regels en limieten zijn in orde.")

                    st.divider()
                    c_h, c_b = st.columns([3, 1])
                    with c_h: st.subheader("Aangepaste Versie")
                    with c_b:
                         st.write("")
                         buffer_edit = maak_excel(edited_df, is_checkbox_data=True)
                         st.download_button("📥 Download Aangepast", buffer_edit, "Rooster_Aangepast.xlsx")

                # =========================================================
                # 📧 MAIL MERGE BESTAND GENEREREN (NIEUW!)
                # =========================================================
                st.divider()
                st.subheader("📧 E-mailen naar muzikanten")
                
                # Check of er een e-mail kolom is
                email_cols = [c for c in df.columns if 'mail' in c.lower()]
                
                if not email_cols:
                    st.info("💡 Tip: Als je in je Excel een kolom 'Email' toevoegt, kan ik een verzendlijst maken.")
                else:
                    st.write("Download hier een lijst die klaar is voor **Word Afdruk Samenvoegen**.")
                    email_col = email_cols[0]
                    
                    if st.button("Genereer Verzendlijst"):
                        mailing_data = []
                        huidige_df = st.session_state['bewerkte_df']
                        unieke_namen = huidige_df['Naam'].unique()
                        
                        for naam in unieke_namen:
                            persoons_rijen = huidige_df[huidige_df['Naam'] == naam]
                            # Mailadres ophalen uit originele DF
                            try:
                                email_adres = df[df['Naam'] == naam][email_col].iloc[0]
                            except:
                                email_adres = "Onbekend"
                            
                            instrumenten_str = ", ".join(persoons_rijen['Instrument'].unique())
                            
                            # Shows verzamelen
                            shows_te_spelen = []
                            for s in shows:
                                if persoons_rijen[s].any(): shows_te_spelen.append(s)
                            
                            mailing_data.append({
                                "Naam": naam,
                                "Email": email_adres,
                                "Instrument": instrumenten_str,
                                "Rooster": ", ".join(shows_te_spelen) if shows_te_spelen else "Geen shows"
                            })
                        
                        mail_df = pd.DataFrame(mailing_data)
                        buffer_mail = io.BytesIO()
                        with pd.ExcelWriter(buffer_mail, engine='xlsxwriter') as writer:
                            mail_df.to_excel(writer, index=False)
                        buffer_mail.seek(0)
                        
                        st.download_button("📥 Download Verzendlijst", buffer_mail, "Verzendlijst.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("Kon geen oplossing vinden. Misschien zijn je regels te streng?")