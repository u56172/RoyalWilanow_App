import os
import sys
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

COLUMNS = ('Data', 'Godzina', 'Imie i Nazwisko', 'Firma', 'Cel')


def get_paths():
    app_dir = os.path.dirname(os.path.abspath(__file__))
    return app_dir, app_dir


def wczytaj_recepcje(plik):
    xls = pd.ExcelFile(plik)
    frames = []

    for arkusz in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=arkusz, usecols='B:F', skiprows=7, header=None, dtype=str)
        df.columns = COLUMNS
        df = df.fillna('').astype(str).apply(lambda x: x.str.strip())
        if 'Firma' in df.columns and (df['Firma'].str.strip() != '').any():
            frames.append(df)
    if len(frames) > 0:
        return pd.concat(frames, ignore_index=True)
    else:
        return pd.DataFrame(columns=COLUMNS)

def grupuj_i_scalaj(sciezki):
    grupy = {'A': [], 'B': [], 'D': [], 'E': [], 'F': []}
    for sciezka in sciezki:
        nazwa = os.path.basename(sciezka).upper()
        for litera in grupy.keys():
            if f'REJESTR_{litera}' in nazwa:
                grupy[litera].append(sciezka)
    dane = {}
    for litera, lista in grupy.items():
        frames = [wczytaj_recepcje(p) for p in lista]
        dane[litera] = pd.concat(frames, ignore_index=True)
    return dane


def przetworz(sciezki):
    dane = grupuj_i_scalaj(sciezki)
    data_A = dane['A'].copy()
    data_B = dane['B'].copy()
    data_D = dane['D'].copy()
    data_E = dane['E'].copy()
    data_F = dane['F'].copy()

    real_management = data_A[data_A['Firma'] == 'real management'].shape[0]
    neo_investments = data_A[data_A['Firma'] == 'neo investments'].shape[0]
    oex = data_A[data_A['Firma'] == 'oex'].shape[0]
    ipos = data_A[data_A['Firma'] == 'ipos'].shape[0]
    ledvance_a = data_A[data_A['Firma'] == 'ledvance'].shape[0]
    budlex_residential = data_A[(data_A['Firma'] == 'budlex residential') | (data_A['Firma'] == 'budlex')].shape[0]
    pm_sport = data_A[data_A['Firma'] == 'pm sport'].shape[0]
    uno_capital = data_A[data_A['Firma'] == 'uno capital'].shape[0]
    cambio = data_A[data_A['Firma'] == 'cambio'].shape[0]
    arrant = data_A[data_A['Firma'] == 'arrant'].shape[0]
    ebs_bud = data_A[(data_A['Firma'] == 'ebs-bud') | (data_A['Firma'] == 'ebs') | (data_A['Firma'] == 'ebs bud')].shape[0]
    bee_creative = data_A[data_A['Firma'] == 'bee creative'].shape[0]
    emitel = data_A[data_A['Firma'] == 'emitel'].shape[0]
    suma_a = real_management + neo_investments + oex + ipos + ledvance_a + budlex_residential + pm_sport + uno_capital + cambio + arrant + ebs_bud + bee_creative + emitel
    capital_park = data_B[(data_B['Firma'] == 'capital park') | (data_B['Firma'] == 'cp')].shape[0]
    refield = data_B[data_B['Firma'] == 'refield'].shape[0]
    spie = data_B[data_B['Firma'] == 'spie'].shape[0]
    dekada = data_B[data_B['Firma'] == 'dekada'].shape[0]
    mjm = data_B[data_B['Firma'] == 'mjm'].shape[0]
    attis = data_B[data_B['Firma'] == 'attis'].shape[0]
    loyalty_point = data_B[(data_B['Firma'] == 'loyalty point') | (data_B['Firma'] == 'lp')].shape[0]
    enata_bread = data_B[(data_B['Firma'] == 'enata bread') | (data_B['Firma'] == 'eb')].shape[0]
    east_management = data_B[(data_B['Firma'] == 'east management') | (data_B['Firma'] == 'em')].shape[0]
    solter = data_B[data_B['Firma'] == 'solter'].shape[0]
    eco_run = data_B[(data_B['Firma'] == 'eco run') | (data_B['Firma'] == 'er')].shape[0]
    suma_b = capital_park + refield + spie + dekada + mjm + attis + loyalty_point + enata_bread + east_management + solter + eco_run
    kingspan = data_D[data_D['Firma'] == 'kingspan'].shape[0]
    essmann = data_D[data_D['Firma'] == 'essmann'].shape[0]
    colt = data_D[data_D['Firma'] == 'colt'].shape[0]
    erbud = data_D[data_D['Firma'] == 'erbud'].shape[0]
    suma_d = kingspan + essmann + colt + erbud
    perfect_gym = data_E[data_E['Firma'] == 'perfect gym'].shape[0]
    health_labs = data_E[(data_E['Firma'] == 'health labs') | (data_E['Firma'] == 'health labs care')].shape[0]
    hilti = data_E[data_E['Firma'] == 'hilti'].shape[0]
    zklaster = data_E[data_E['Firma'] == 'zklaster'].shape[0]
    gepol_solar = data_E[data_E['Firma'] == 'gepol solar'].shape[0]
    rsx = data_E[data_E['Firma'] == 'rsx'].shape[0]
    gedeon_richter = data_E[data_E['Firma'] == 'gedeon richter'].shape[0]
    ledvance_e = data_E[data_E['Firma'] == 'ledvance'].shape[0]
    mjm_services = data_E[data_E['Firma'] == 'mjm services'].shape[0]
    mjm_brokers = data_E[data_E['Firma'] == 'mjm brokers'].shape[0]
    mhc_mobility_hitachi = data_E[(data_E['Firma'] == 'mhc mobility') | (data_E['Firma'] == 'hitachi')].shape[0]
    suma_e = perfect_gym + health_labs + hilti + zklaster + gepol_solar + rsx + gedeon_richter + ledvance_e + mjm_services + mjm_brokers + mhc_mobility_hitachi
    lindt = data_F[data_F['Firma'] == 'lindt'].shape[0]
    pirelli_prometeon = data_F[(data_F['Firma'] == 'pirelli') | (data_F['Firma'] == 'prometeon')].shape[0]
    ajinomoto = data_F[data_F['Firma'] == 'ajinomoto'].shape[0]
    neo_energy = data_F[data_F['Firma'] == 'neo energy'].shape[0]
    suma_f = lindt + pirelli_prometeon + ajinomoto + neo_energy
    recepcje_counts = {'A': suma_a, 'B': suma_b, 'D': suma_d, 'E': suma_e, 'F': suma_f}
    suma_gosci = int(suma_a + suma_b + suma_d + suma_e + suma_f)

    all_data = pd.concat([data_A, data_B, data_D, data_E, data_F], ignore_index=True)
    firmy_counts = all_data['Firma'].str.strip().value_counts()
    most_amount_company = firmy_counts.idxmax()
    most_amount_company_count = int(firmy_counts.max())

    return {
        'dane': dane,
        'suma': recepcje_counts,
        'suma_gosci': suma_gosci,
        'top_firma': most_amount_company,
        'top_firma_count': most_amount_company_count,
        'calosc': {
            'A': (real_management, neo_investments, oex + ipos, ledvance_a, budlex_residential, uno_capital, cambio,
                  pm_sport, bee_creative, ebs_bud, arrant, emitel, suma_a),
            'B': (capital_park, refield, spie, dekada, mjm + attis, loyalty_point + enata_bread,
                  east_management + solter, eco_run, suma_b),
            'D': (kingspan + essmann + colt, erbud, suma_d),
            'E': (perfect_gym + health_labs, hilti, mjm_services + mjm_brokers, ledvance_e,
                  zklaster + gepol_solar + rsx + gedeon_richter, mhc_mobility_hitachi, suma_e),
            'F': (lindt, pirelli_prometeon, ajinomoto, neo_energy, suma_f)
        }
    }


def zapisz_pliki(wynik, out_dir):
    def p(name): return os.path.join(out_dir, name)
    os.makedirs(out_dir, exist_ok=True)
    a = wynik['calosc']['A']
    b = wynik['calosc']['B']
    d = wynik['calosc']['D']
    e = wynik['calosc']['E']
    f = wynik['calosc']['F']
    suma = wynik['suma']
    suma_gosci = wynik['suma_gosci']
    with open(p('recepcja_a.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja A\nDo Real Management: {a[0]}\nDo Neo Investments: {a[1]}\nDo OEX/IPOS: {a[2]}\nDo Ledvance: {a[3]}\nDo Budlex: {a[4]}\nDo Uno Capital: {a[5]}\nDo Cambio: {a[6]}\nDo PM Sport: {a[7]}\nDo Bee Creative: {a[8]}\nDo EBS-BUD: {a[9]}\nDo Arrant: {a[10]}\nDo Emitel: {a[11]}\nŁĄCZNIE: {a[12]}\n")
    with open(p('recepcja_b.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja B\nDo Capital Park: {b[0]}\nDo Refield: {b[1]}\nDo SPIE: {b[2]}\nDo Dekada: {b[3]}\nDo MJM/Attis: {b[4]}\nDo Loyalty Point/Enata Bread: {b[5]}\nDo East Management/Solter: {b[6]}\nDo Eco Run: {b[7]}\nŁĄCZNIE: {b[8]}\n")
    with open(p('recepcja_d.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja D\nDo Kingspan/Essmann/Colt: {d[0]}\nDo Erbud: {d[1]}\nŁĄCZNIE: {d[2]}\n")
    with open(p('recepcja_e.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja E\nDo Perfect Gym/Health Labs Care: {e[0]}\nDo Hilti: {e[1]}\nDo MJM Services/Brokers: {e[2]}\nDo Ledvance: {e[3]}\nDo Zklaster/Gepol Solar/RSX: {e[4]}\nDo MHC Mobility: {e[5]}\nŁĄCZNIE: {e[6]}\n")
    with open(p('recepcja_f.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja F\nDo Lindt: {f[0]}\nDo Pirelli i Prometeon: {f[1]}\nDo Ajinomoto: {f[2]}\nDo Neo Energy: {f[3]}\nŁĄCZNIE: {f[4]}\n")
    with open(p('recepcje_suma.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja A: {suma['A']} osób\nRecepcja B: {suma['B']} osób\nRecepcja D: {suma['D']} osób\nRecepcja E: {suma['E']} osób\nRecepcja F: {suma['F']} osób\nŁĄCZNIE: {suma_gosci} osób\n")
    perc = {k: round(suma[k] / suma_gosci * 100, 1) if suma_gosci else 0 for k in suma}
    with open(p('recepcje_perc.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Recepcja A: {perc['A']}%\nRecepcja B: {perc['B']}%\nRecepcja D: {perc['D']}%\nRecepcja E: {perc['E']}%\nRecepcja F: {perc['F']}%\n")
    most_amount_reception = max(suma, key=suma.get) if suma_gosci else None
    most_amount_reception_count = suma[most_amount_reception] if most_amount_reception else 0
    with open(p('recepcja_end.txt'), 'w', encoding='utf-8') as ftxt:
        ftxt.write(f"Liczba gości w miesiącu: {suma_gosci}\nNajwięcej gości w m-cu - recepcja: {most_amount_reception} ({most_amount_reception_count} osób)\n" + \
            (f"Najwięcej gości w m-cu - firma: {wynik['top_firma']} ({wynik['top_firma_count']} osób)\n" if wynik['top_firma'] else ""))
    plt.figure(figsize=(8, 6))
    plt.bar(list(suma.keys()), list(suma.values()))
    plt.title('Liczba gości Royal Wilanów [os.]')
    plt.xlabel('Recepcja')
    plt.ylabel('Liczba gości [os.]')

    for i, v in enumerate(suma.values()):
        plt.text(i, v + 1, str(v), ha='center', va='bottom')
    plt.tight_layout()
    plt.savefig(p('wykres_slupkowy.png'), dpi=150)
    plt.close()
    plt.figure(figsize=(7, 7))
    if suma_gosci:
        plt.pie(list(suma.values()), labels=[f"Recepcja {k}" for k in suma.keys()], autopct='%1.1f%%', startangle=90,
                counterclock=False)
    plt.title('Udział procentowy w ogólnej liczbie gości [%]')
    plt.tight_layout()
    plt.savefig(p('wykres_kolowy.png'), dpi=150)
    plt.close()