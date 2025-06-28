import os
import pandas as pd
from datetime import datetime

STATUS_ONLINE = 'on'
STATUS_OFFLINE = 'off'
TABLE_STYLE = 'Table Style Medium 2'

banner = '''

               ▐▄• ▄ ▄▄▄▄▄▄▄▄   ▄▄▄·  ▄▄· ▄▄▄▄▄
                █▌█▌▪•██  ▀▄ █·▐█ ▀█ ▐█ ▌▪•██  
                ·██·  ▐█.▪▐▀▀▄ ▄█▀▀█ ██ ▄▄ ▐█.▪
               ▪▐█·█▌ ▐█▌·▐█•█▌▐█ ▪▐▌▐███▌ ▐█▌·
               •▀▀ ▀▀ ▀▀▀ .▀  ▀ ▀  ▀ ·▀▀▀  ▀▀▀ 
                                                                           
'''

banner1 = '''
      
      $$\   $$\ $$$$$$$$\ $$$$$$$\   $$$$$$\   $$$$$$\ $$$$$$$$\ 
      $$ |  $$ |\__$$  __|$$  __$$\ $$  __$$\ $$  __$$\\__$$  __|
      \$$\ $$  |   $$ |   $$ |  $$ |$$ /  $$ |$$ /  \__|  $$ |   
       \$$$$  /    $$ |   $$$$$$$  |$$$$$$$$ |$$ |        $$ |   
       $$  $$<     $$ |   $$  __$$< $$  __$$ |$$ |        $$ |   
      $$  /\$$\    $$ |   $$ |  $$ |$$ |  $$ |$$ |  $$\   $$ |   
      $$ /  $$ |   $$ |   $$ |  $$ |$$ |  $$ |\$$$$$$  |  $$ |   
      \__|  \__|   \__|   \__|  \__|\__|  \__| \______/   \__|   
                                                                                                                                                                                                                                     
'''

banner2 = '''
      
      ▒██   ██▒▄▄▄█████▓ ██▀███   ▄▄▄       ▄████▄  ▄▄▄█████▓
      ▒▒ █ █ ▒░▓  ██▒ ▓▒▓██ ▒ ██▒▒████▄    ▒██▀ ▀█  ▓  ██▒ ▓▒
      ░░  █   ░▒ ▓██░ ▒░▓██ ░▄█ ▒▒██  ▀█▄  ▒▓█    ▄ ▒ ▓██░ ▒░
       ░ █ █ ▒ ░ ▓██▓ ░ ▒██▀▀█▄  ░██▄▄▄▄██ ▒▓▓▄ ▄██▒░ ▓██▓ ░ 
      ▒██▒ ▒██▒  ▒██▒ ░ ░██▓ ▒██▒ ▓█   ▓██▒▒ ▓███▀ ░  ▒██▒ ░ 
      ▒▒ ░ ░▓ ░  ▒ ░░   ░ ▒▓ ░▒▓░ ▒▒   ▓▒█░░ ░▒ ▒  ░  ▒ ░░   
      ░░   ░▒ ░    ░      ░▒ ░ ▒░  ▒   ▒▒ ░  ░  ▒       ░    
       ░    ░    ░        ░░   ░   ░   ▒   ░          ░      
       ░    ░              ░           ░  ░░ ░               
                                           ░                 
'''

banners = [banner, banner1, banner2]

def garantir_colunas(df, colunas, valor_default=''):
    for col in colunas:
        if col not in df.columns:
            df[col] = valor_default
    ####  DEBUG 
    #print("\n[DEBUG] Colunas do DataFrame:")
    #for col in df.columns:
    #    print(f"- {col}")
    return df


def formatar_aba(writer, df, aba):
    workbook = writer.book
    worksheet = writer.sheets[aba]
    max_row, max_col = df.shape
    column_settings = [{'header': col} for col in df.columns]

    if max_row > 0:
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': column_settings,
            'style': TABLE_STYLE
        })

    for i, col in enumerate(df.columns):
        max_len = max(df[col].astype(str).map(len).max(), len(col))
        worksheet.set_column(i, i, max_len * 1.2)

def formatar_excel(df, output_file, aba):
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=aba, index=False)
            formatar_aba(writer, df, aba)
    except Exception as e:
        fallback = output_file.replace('.xlsx', '.csv')
        df.to_csv(fallback, index=False)
        print(f"[\033[1;31m!\033[m] Erro ao salvar Excel ({e}). CSV criado: {fallback}")

def clean_field(val):
    if isinstance(val, list):
        if not val:
            return ''
        if all(isinstance(i, str) for i in val):
            return ', '.join(val)
        if all(isinstance(i, dict) for i in val):
            return ', '.join(filter(None, [i.get('name', '') for i in val]))
        return str(val[0])
    if isinstance(val, dict):
        return val.get('name', str(val))
    return val

def validar_credenciais(tio):
    try:
        session = tio.session.details()
        print(f"[\033[1;32m✓\033[m] Conectado como: \033[1m{session['username']}\033[0m")
        return True
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Falha na autenticação: {e}")
        return False

def groups_to_str(df):
    df = df.copy()
    df['groups'] = df['groups'].apply(
        lambda v: ', '.join([g.get('name', '') for g in v]) if isinstance(v, list) else ''
    )
    return df


def prepare_df(dff):
    dff = dff.copy()
    if 'core_version' in dff.columns:
        dff['version_agent'] = dff['core_version'].fillna('N/A').replace('', 'N/A')
    else:
        dff['version_agent'] = 'N/A'

    def ts(t):
        try:
            return datetime.fromtimestamp(int(t)).strftime('%Y-%m-%d %H:%M:%S') if t else ''
        except:
            return ''

    for col in ['linked_on', 'last_scanned', 'last_connect']:
        if col in dff.columns:
            dff[col] = dff[col].apply(ts)
        else:
            dff[col] = ''

    dff['status'] = dff['status'].apply(lambda x: 'Online' if str(x).lower() == STATUS_ONLINE else 'Offline')

    cols = [
        'status', 'name', 'ip', 'platform', 'version_agent', 'groups',
        'linked_on', 'last_scanned', 'last_connect', 'uuid'
    ]
    dff = garantir_colunas(dff, cols)
    dff = dff[cols]

    dff.columns = [
        'Status', 'Agent Name', 'IP Address', 'Platform', 'Version', 'Groups',
        'Linked On', 'Last Scanned', 'Last Connect', 'Agent UUID'
    ]
    return dff

def is_group_empty(x):
    if isinstance(x, list):
        return len(x) == 0
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return True
    if isinstance(x, str) and x.strip() == '':
        return True
    return False

def padronizar_excel_agents(caminho_arquivo):
    try:
        print(f"[\033[1m1/4\033[0m] Lendo arquivo: {caminho_arquivo}")
        if caminho_arquivo.endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
        elif caminho_arquivo.endswith('.xlsx'):
            df = pd.read_excel(caminho_arquivo)
        else:
            print("[\033[1;31m!\033[m] Formato de arquivo não suportado. Use .csv ou .xlsx")
            return

        df = df.applymap(clean_field)

        if 'Agent Name' in df.columns:
            df = df.rename(columns={
                'Status': 'status',
                'Agent Name': 'name',
                'IP Address': 'ip',
                'Platform': 'platform',
                'Version': 'core_version',
                'Groups': 'groups',
                'Linked On': 'linked_on',
                'Last Scanned': 'last_scanned',
                'Last Connect': 'last_connect',
                'Agent UUID': 'uuid'
            })

        colunas_padrao = [
            'status', 'name', 'ip', 'platform', 'core_version', 'groups',
            'linked_on', 'last_scanned', 'last_connect', 'uuid'
        ]
        df = garantir_colunas(df, colunas_padrao)

        df['status'] = df['status'].apply(lambda x: 'Online' if str(x).lower() == STATUS_ONLINE else 'Offline')
        df['core_version'] = df['core_version'].fillna('N/A').replace('', 'N/A')

        def ts(t):
            try:
                if pd.isna(t) or t == '':
                    return ''
                if str(t).isdigit():
                    return datetime.fromtimestamp(int(t)).strftime('%Y-%m-%d %H:%M:%S')
                t_str = str(t)
                if t_str.endswith('Z'):
                    t_str = t_str[:-1]  # Remove 'Z' final
                dt = datetime.strptime(t_str[:19], '%Y-%m-%dT%H:%M:%S')
                return dt.strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                return str(t)

        for col in ['linked_on', 'last_scanned', 'last_connect']:
            df[col] = df[col].apply(ts)

        df['groups'] = df['groups'].apply(lambda v: ', '.join([g.get('name', '') for g in v]) if isinstance(v, list) else str(v))

        df_final = df[colunas_padrao].copy()
        df_final.columns = [
            'Status', 'Agent Name', 'IP Address', 'Platform', 'Version', 'Groups',
            'Linked On', 'Last Scanned', 'Last Connect', 'Agent UUID'
        ]

        output_folder = os.path.join(os.getcwd(), 'Padronizados')
        os.makedirs(output_folder, exist_ok=True)

        nome_arquivo = os.path.basename(caminho_arquivo).replace('.csv', '').replace('.xlsx', '')
        output_file = os.path.join(output_folder, f'{nome_arquivo}_padronizado.xlsx')

        formatar_excel(df_final, output_file, 'Agents')
        print(f"[\033[1;32m✓\033[m] Arquivo padronizado salvo em: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao padronizar arquivo: {e}")
        traceback.print_exc()


def loading_bar(iteration, total, prefix='', suffix='', length=50):
    import sys

    percent = f"{100 * (iteration / float(total)):.0f}"
    filled_length = int(length * iteration // total)
    
    bar = f"\033[1;31m{'―' * filled_length}\033[1;30m{'-' * (length - filled_length)}\033[0m" 
    output = f'\r{prefix} |{bar}| {percent}% {suffix.ljust(25)}'

    if iteration == total:
        output += '\n'

    sys.stdout.write(output)
    sys.stdout.flush()
