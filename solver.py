import pulp
import pandas as pd

def run_solver(df, limits_per_instrument, extra_regels):

    df = df.dropna(subset=['Naam'])
    
    # 1. DATA VOORBEREIDEN
    # Unieke ID maken voor elke rij (voor het geval Jan twee instrumenten speelt)
    df['Resource_ID'] = df['Naam'].astype(str) + " (" + df['Instrument'].astype(str) + ")"
    
    resources = df['Resource_ID'].tolist()
    unieke_personen = df['Naam'].unique().tolist()
    instrumenten = df['Instrument'].unique().tolist()
    
    # Bepaal de show kolommen (alles wat geen standaard kolom is)
    vaste_kolommen = ['Naam', 'Instrument', 'Wens', 'Label', 'Resource_ID']
    momenten = [c for c in df.columns if c not in vaste_kolommen]

    # Lookups
    res_instr_map = df.set_index('Resource_ID')['Instrument'].to_dict()
    res_naam_map = df.set_index('Resource_ID')['Naam'].to_dict()
    # Wens (eerste wens die we vinden per persoon)
    persoon_wens_map = df.groupby('Naam')['Wens'].first().to_dict()

    # 2. MODEL OPBOUWEN
    prob = pulp.LpProblem("Orkest_Planner", pulp.LpMaximize)

    # Variabelen: x[resource, moment]
    x = pulp.LpVariable.dicts("inzet", 
                              ((r, m) for r in resources for m in momenten), 
                              cat='Binary')

    # 3. DOELFUNCTIE
    score = 0
    for r in resources:
        for m in momenten:
            val = df[df['Resource_ID'] == r][m].values[0]
            try:
                cijfer = int(val)
            except:
                cijfer = 0 # Valback als het geen getal is
            
            waarde = 0
            if cijfer == 3: waarde = 10
            elif cijfer == 2: waarde = 5
            elif cijfer == 1: waarde = -100
            elif cijfer == 0: waarde = -10000 
            
            score += x[r, m] * waarde

    prob += score

    # 4. CONSTRAINTS
    
    # A. Bezetting per Instrument (Sidebar Limieten)
    for m in momenten:
        for instr in instrumenten:
            if instr in limits_per_instrument:
                max_aantal = limits_per_instrument[instr]
                # Tel alle resources van dit instrument
                aantal_ingepland = pulp.lpSum([x[r, m] for r in resources if res_instr_map[r] == instr])
                prob += aantal_ingepland <= max_aantal

    # B. UNICITEIT (1 persoon kan maar 1 ding tegelijk doen)
    for p in unieke_personen:
        mijn_rollen = [r for r in resources if res_naam_map[r] == p]
        for m in momenten:
            prob += pulp.lpSum([x[r, m] for r in mijn_rollen]) <= 1

    # C. WENS (Totaal aantal shows)
    for p in unieke_personen:
        wens_val = persoon_wens_map.get(p, "-")
        try:
            gewenst = int(wens_val)
            if gewenst > 0:
                mijn_rollen = [r for r in resources if res_naam_map[r] == p]
                totaal_ingepland = pulp.lpSum([x[r, m] for r in mijn_rollen for m in momenten])
                prob += totaal_ingepland >= gewenst - 1
                prob += totaal_ingepland <= gewenst + 1
        except:
            pass # Geen getal ingevuld, negeer

    # D. EXTRA REGELS
    for regel in extra_regels:
        try:
            # 1. Conflict (Niet Samen)
            if regel['type'] == 'conflict':
                p1, p2 = regel['p1'], regel['p2']
                if p1 in unieke_personen and p2 in unieke_personen:
                    rollen_p1 = [r for r in resources if res_naam_map[r] == p1]
                    rollen_p2 = [r for r in resources if res_naam_map[r] == p2]
                    for m in momenten:
                        prob += pulp.lpSum([x[r, m] for r in rollen_p1]) + pulp.lpSum([x[r, m] for r in rollen_p2]) <= 1
                        
            # 2. Samen
            elif regel['type'] == 'samen':
                p1, p2 = regel['p1'], regel['p2']
                if p1 in unieke_personen and p2 in unieke_personen:
                    rollen_p1 = [r for r in resources if res_naam_map[r] == p1]
                    rollen_p2 = [r for r in resources if res_naam_map[r] == p2]
                    for m in momenten:
                        prob += pulp.lpSum([x[r, m] for r in rollen_p1]) == pulp.lpSum([x[r, m] for r in rollen_p2])

            # 3. Must All
            elif regel['type'] == 'must_all':
                p1 = regel['p1']
                if p1 in unieke_personen:
                    rollen_p1 = [r for r in resources if res_naam_map[r] == p1]
                    for m in momenten:
                        prob += pulp.lpSum([x[r, m] for r in rollen_p1]) == 1

            # 4. Force Show
            elif regel['type'] == 'force_show':
                p1_naam = regel['p1']
                target_show = regel['show']
                if target_show in momenten:
                    p1_ids = [r for r in resources if res_naam_map[r] == p1_naam]
                    prob += pulp.lpSum([x[pid, target_show] for pid in p1_ids]) == 1

            # 5. Minimaal aantal shows
            elif regel['type'] == 'min_shows':
                p1_naam = regel['p1']
                min_aantal = regel['count']
                p1_ids = [r for r in resources if res_naam_map[r] == p1_naam]
                
                # Verzamel alle inzet variabelen
                alle_inzet = []
                for pid in p1_ids:
                    for show in momenten:
                        alle_inzet.append(x[pid, show])
                prob += pulp.lpSum(alle_inzet) >= min_aantal
        except Exception as e:
            print(f"Fout bij regel {regel}: {e}")

    # 5. OPLOSSEN
    prob.solve()
    status = pulp.LpStatus[prob.status]
    
    # 6. RESULTAAT
    rooster_data = []
    if status == "Optimal":
        for r in resources:
            row = {
                'Naam': res_naam_map[r],
                'Instrument': res_instr_map[r]
            }
            totaal = 0
            for m in momenten:
                if pulp.value(x[r, m]) == 1:
                    row[m] = "✅"
                    totaal += 1
                else:
                    row[m] = "."
            row['Totaal'] = totaal
            rooster_data.append(row)
    
    return status, pd.DataFrame(rooster_data)