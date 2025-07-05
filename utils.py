import os
import pandas as pd
from datetime import datetime
from tenable.io import TenableIO 
import json
import csv
import arrow
import sys
import re
import shutil
import requests 

from config import STATUS_ONLINE, STATUS_OFFLINE, TABLE_STYLE, PLUGINS_SOFTWARE_INVENTORY, \
    EXPLOITDB_EXPLOITS_CSV_URL, EXPLOITDB_SHELLCODES_CSV_URL, \
    EXPLOITDB_EXPLOITS_LOCAL_CSV_PATH, EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH, \
    DELETE_EXPLOITDB_CSV_AFTER_USE, banners 

_exploitdb_data = None


def garantir_colunas(df: pd.DataFrame, colunas: list, valor_default: any = '') -> pd.DataFrame:
    for col in colunas:
        if col not in df.columns:
            df[col] = valor_default
    return df

def formatar_aba(writer: pd.ExcelWriter, df: pd.DataFrame, aba: str):
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

def formatar_excel_output(data: pd.DataFrame | dict[str, pd.DataFrame], output_file: str, sheet_name: str = 'Sheet1'):
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            if isinstance(data, pd.DataFrame):
                data.to_excel(writer, sheet_name=sheet_name, index=False)
                formatar_aba(writer, data, sheet_name)
            elif isinstance(data, dict):
                for current_sheet_name, df_to_save in data.items():
                    df_to_save.to_excel(writer, sheet_name=current_sheet_name, index=False)
                    formatar_aba(writer, df_to_save, current_sheet_name)
            else:
                raise TypeError("O parâmetro 'data' deve ser um pandas.DataFrame ou um dicionário de pandas.DataFrame.")

    except Exception as e:
        if isinstance(data, pd.DataFrame):
            fallback = output_file.replace('.xlsx', '.csv')
            data.to_csv(fallback, index=False)
            print(f"[\033[1;31m!\033[m] Erro ao salvar Excel ({e}). CSV criado: {fallback}")
        else:
            print(f"[\033[1;31m!\033[m] Erro ao salvar Excel com múltiplas abas ({e}).")

def clean_field(val: any) -> str:
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

def validar_credenciais(tio: TenableIO) -> bool:
    try:
        session = tio.session.details()
        print(f"[\033[1;32m✓\033[m] Conectado como: \033[1m{session['username']}\033[0m")
        return True
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Falha na autenticação: {e}")
        return False

def groups_to_str(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['groups'] = df['groups'].apply(lambda v: v if isinstance(v, list) else [])
    
    df['groups'] = df['groups'].apply(
        lambda v: ', '.join([g.get('name', '') for g in v if isinstance(g, dict) and g.get('name')])
                   if v 
                   else 'N/A'
    )
    return df

def convert_timestamp_to_datetime_str(t: any) -> str:
    if pd.isna(t) or t == '':
        return ''
    try:
        if isinstance(t, (int, float)): 
            return datetime.fromtimestamp(int(t)).strftime('%Y-%m-%d %H:%M:%S')
        
        t_str = str(t)
        dt_obj = pd.to_datetime(t_str, errors='coerce', utc=True)
        if pd.isna(dt_obj):
            if t_str.endswith('Z'):
                t_str = t_str[:-1] 
            dt_obj = datetime.strptime(t_str[:19], '%Y-%m-%dT%H:%M:%S') 
        
        return dt_obj.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return str(t)

def prepare_df(dff: pd.DataFrame) -> pd.DataFrame:
    dff = dff.copy()
    if 'core_version' in dff.columns:
        dff['version_agent'] = dff['core_version'].fillna('N/A').replace('', 'N/A')
    else:
        dff['version_agent'] = 'N/A'

    for col in ['linked_on', 'last_scanned', 'last_connect']:
        if col in dff.columns:
            dff[col] = dff[col].apply(convert_timestamp_to_datetime_str)
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

def is_group_empty(x: any) -> bool:

    if isinstance(x, list):
        return len(x) == 0
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return True
    if isinstance(x, str) and x.strip() == '':
        return True
    return False

def padronizar_excel_agents(caminho_arquivo: str):
    try:
        print(f"[\033[1m1/4\033[0m] Lendo arquivo: {caminho_arquivo}")
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
        elif caminho_arquivo.lower().endswith('.xlsx'):
            df = pd.read_excel(caminho_arquivo)
        else:
            print("[\033[1;31m!\033[m] Formato de arquivo não suportado. Use .csv ou .xlsx")
            return
        for col in df.columns:
            df[col] = df[col].apply(clean_field)


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

        for col in ['linked_on', 'last_scanned', 'last_connect']:
            df[col] = df[col].apply(convert_timestamp_to_datetime_str)
        
        df['groups'] = df['groups'].astype(str) 

        df_final = df[colunas_padrao].copy()
        df_final.columns = [
            'Status', 'Agent Name', 'IP Address', 'Platform', 'Version', 'Groups',
            'Linked On', 'Last Scanned', 'Last Connect', 'Agent UUID'
        ]

        output_folder = os.path.join(os.getcwd(), 'Padronizados')
        os.makedirs(output_folder, exist_ok=True)

        nome_arquivo = os.path.basename(caminho_arquivo).rsplit('.', 1)[0]
        output_file = os.path.join(output_folder, f'{nome_arquivo}_padronizado.xlsx')

        formatar_excel_output(df_final, output_file, 'Agents')
        print(f"[\033[1;32m✓\033[m] Arquivo padronizado salvo em: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao padronizar arquivo: {e}")
        traceback.print_exc()

def loading_bar(iteration: int, total: int, prefix: str = '', suffix: str = '', length: int = 50):
    import sys
    try:
        terminal_width = os.get_terminal_size().columns
    except OSError:
        terminal_width = 80 

    # Garante que total não é zero para evitar ZeroDivisionError
    if total == 0:
        percent = "0"
        filled_length = 0
    else:
        percent = f"{100 * (iteration / float(total)):.0f}"
        filled_length = int(length * iteration // total)
    
    bar = f"\033[1;31m{'―' * filled_length}\033[1;30m{'-' * (length - filled_length)}\033[0m" 
    
    current_line_length = len(prefix) + len(" |") + length + len("| ") + len(percent) + len("% ")
    max_suffix_len = terminal_width - current_line_length - 2 
    effective_suffix_len = max(0, max_suffix_len)

    if len(suffix) > effective_suffix_len:
        if effective_suffix_len > 3:
            display_suffix = suffix[:effective_suffix_len - 3] + '...'
        else:
            display_suffix = suffix[:effective_suffix_len]
    else:
        display_suffix = suffix.ljust(effective_suffix_len)

    output = f'\r{prefix} |{bar}| {percent}% {display_suffix}'
    
    sys.stdout.write(output.ljust(terminal_width)) 
    
    if iteration == total:
        sys.stdout.write('\n') 
    
    sys.stdout.flush()

def inventario_software(ACCESS_KEY: str, SECRET_KEY: str, output_folder: str, cliente: str):
    try:
        total_steps = 8
        current_step = 1

        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Inicializando Tenable.io', length=50)
        tio = TenableIO(ACCESS_KEY, SECRET_KEY)
        current_step += 1

        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Coletando Findings (últimos 90 dias)', length=50)
        findings_raw = list(tio.exports.vulns(last_found=int(arrow.now().shift(days=-90).timestamp())))
        current_step += 1
        
        all_software_data = []
        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix=f'Processando {len(findings_raw)} Findings', length=50)
        
        for finding in findings_raw:
            plugin_id = finding.get("plugin", {}).get("id")
            if plugin_id in PLUGINS_SOFTWARE_INVENTORY: 
                hostname = finding.get("asset", {}).get("fqdn", "N/A")
                ip_address = finding.get("asset", {}).get("ipv4", "N/A")
                os_type = finding.get("asset", {}).get("operating_system", "N/A")
                plugin_output = finding.get("output", "") 

                software_list_raw = plugin_output.split('\n')
                for software_line in software_list_raw:
                    cleaned_software = software_line.strip()
                    if cleaned_software and not any(
                        cleaned_software.lower().startswith(header_prefix.lower()) for header_prefix in [
                            "the following software are installed on the remote host :",
                            "here is the list of packages installed on the remote",
                            "installed software:" 
                        ]
                    ):
                        all_software_data.append({
                            "Hostname": hostname,
                            "IP Address": ip_address,
                            "OS": os_type,
                            "Software": cleaned_software,
                            "Plugin ID": plugin_id 
                        })
        current_step += 1

        if not all_software_data:
            print("\n[\033[1;33m!\033[m] Nenhum software encontrado para os plugins especificados nos últimos 90 dias.")
            return

        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Criando DataFrames', length=50)
        df_software_all = pd.DataFrame(all_software_data)
        current_step += 1 #4
        
        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Separando e Garantindo Colunas', length=50)
        colunas_final_output = ["Hostname", "IP Address", "OS", "Software"]

        df_software_all['Plugin ID'] = pd.to_numeric(df_software_all['Plugin ID'], errors='coerce')
        
        df_windows = df_software_all[df_software_all['Plugin ID'] == 20811].drop(columns=['Plugin ID'])
        df_linux = df_software_all[df_software_all['Plugin ID'] == 22869].drop(columns=['Plugin ID'])
        df_todos = df_software_all.drop(columns=['Plugin ID']) 

        df_windows = garantir_colunas(df_windows, colunas_final_output)
        df_linux = garantir_colunas(df_linux, colunas_final_output)
        df_todos = garantir_colunas(df_todos, colunas_final_output)
        current_step += 1 #5

        output_file = os.path.join(output_folder, f'Tenable_Software_Inventory_{cliente}.xlsx')
        
        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Organizando DataFrames para Exportação', length=50)
        dfs_to_export = {}
        
        if not df_todos.empty:
            dfs_to_export['Todos'] = df_todos
        if not df_windows.empty:
            dfs_to_export[PLUGINS_SOFTWARE_INVENTORY[20811]] = df_windows
        if not df_linux.empty:
            dfs_to_export[PLUGINS_SOFTWARE_INVENTORY[22869]] = df_linux
        
        if not dfs_to_export:
            print("\n[\033[1;33m!\033[m] Nenhum dado de software para Windows, Linux ou a aba 'Todos' para exportar.")
            return
        current_step += 1 #6

        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Salvando Arquivo Excel', length=50)
        formatar_excel_output(dfs_to_export, output_file)
        current_step += 1 #7
        
        loading_bar(current_step, total_steps, prefix='Inventário de Software:', suffix='Concluído', length=50)
        print(f"\n[\033[1;32m✓\033[m] Inventário de software salvo em: {output_file}")

    except Exception as e:
        import traceback
        print(f"\n[\033[1;31m!\033[m] Erro ao gerar inventário de software: {e}")
        traceback.print_exc()


def download_file(url: str, local_path: str, description: str, overall_prefix: str, current_overall_iteration: int, total_overall_iterations: int) -> bool:

    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()

        total_size = int(response.headers.get('content-length', 0))
        downloaded_size = 0
        block_size = 8192

        with open(local_path, 'wb') as f:
            for data in response.iter_content(block_size):
                downloaded_size += len(data)
                f.write(data)
                

                percent_this_download = int(100 * (downloaded_size / total_size)) if total_size > 0 else 0
                

                loading_bar(percent_this_download, 100, prefix=f" - Baixando {description}:",
                            suffix=f"{downloaded_size / (1024*1024):.1f}MB / {total_size / (1024*1024):.1f}MB", length=50)
        
        sys.stdout.write('\n') 
        sys.stdout.flush() 

        return True

    except requests.exceptions.RequestException as e:
        sys.stdout.write('\n') 
        print(f"[\033[1;31m!\033[m] Erro ao baixar {description}: {e}")
        import traceback; traceback.print_exc()
        return False
    except Exception as e:
        sys.stdout.write('\n') 
        print(f"[\033[1;31m!\033[m] Erro inesperado no download de {description}: {e}")
        import traceback; traceback.print_exc()
        return False


def load_exploitdb_data(current_overall_step: int, total_overall_steps: int, overall_prefix: str) -> pd.DataFrame:
    global _exploitdb_data
    if _exploitdb_data is not None:
        loading_bar(current_overall_step, total_overall_steps, overall_prefix, suffix="Exploit-DB carregado do cache.", length=50)
        return _exploitdb_data


    num_sub_steps = 3
    progress_per_sub_step = 1.0 / num_sub_steps 



    loading_bar(current_overall_step, total_overall_steps, overall_prefix, suffix=" Baixando Exploit-DB Exploits CSV...", length=50)
    if not download_file(EXPLOITDB_EXPLOITS_CSV_URL, EXPLOITDB_EXPLOITS_LOCAL_CSV_PATH, 
                         "Exploit-DB Exploits CSV", current_overall_step, total_overall_steps, overall_prefix):
        return pd.DataFrame() 

    shellcodes_df = pd.DataFrame()
    if EXPLOITDB_SHELLCODES_CSV_URL and EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH:
        loading_bar(current_overall_step + int(1 * progress_per_sub_step), total_overall_steps, overall_prefix, suffix=" Baixando Exploit-DB Shellcodes CSV...", length=50)
        if download_file(EXPLOITDB_SHELLCODES_CSV_URL, EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH, 
                         "Exploit-DB Shellcodes CSV", current_overall_step + int(1 * progress_per_sub_step), total_overall_steps, overall_prefix):
            try:
                shellcodes_df = pd.read_csv(EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH, encoding='utf-8', on_bad_lines='skip')
                if shellcodes_df.empty or len(shellcodes_df.columns) == 0:
                    raise pd.errors.EmptyDataError("Shellcodes CSV is empty or has no columns.")

                shellcodes_df = shellcodes_df[['id', 'codes', 'description', 'platform', 'author', 'date_published']].copy()
                shellcodes_df = shellcodes_df.rename(columns={'id': 'exploit_id', 'codes': 'cve', 'description': 'exploit_description', 'platform': 'exploit_platform', 'author': 'exploit_author', 'date_published': 'exploit_date_published'})
                shellcodes_df['cve'] = shellcodes_df['cve'].astype(str).apply(lambda x: [c.strip().upper() for c in x.split(',') if c.strip().upper().startswith('CVE-')])
                shellcodes_df = shellcodes_df.explode('cve')
                shellcodes_df = shellcodes_df[shellcodes_df['cve'].notna() & (shellcodes_df['cve'] != '')]
                shellcodes_df['exploit_db_url'] = shellcodes_df['exploit_id'].apply(lambda x: f"https://www.exploit-db.com/shellcodes/{int(x)}" if pd.notna(x) else '')
            except pd.errors.EmptyDataError:
                print(f"[\033[1;33m!\033[m] Aviso: Shellcodes CSV está vazio ou sem colunas. Ignorando...")
                shellcodes_df = pd.DataFrame() 
            except Exception as e:
                print(f"[\033[1;31m!\033[m] Erro ao carregar/processar Shellcodes CSV: {e}")
                import traceback; traceback.print_exc()
                shellcodes_df = pd.DataFrame()


    loading_bar(current_overall_step + int(2 * progress_per_sub_step), total_overall_steps, overall_prefix, suffix="Processando dados Exploit-DB em memória...", length=50)

    try:
        df_exploitdb = pd.read_csv(EXPLOITDB_EXPLOITS_LOCAL_CSV_PATH, encoding='utf-8', on_bad_lines='skip')
        
        if df_exploitdb.empty or len(df_exploitdb.columns) == 0:
            raise pd.errors.EmptyDataError("Exploits CSV está vazio ou sem colunas. Não é possível continuar.")

        df_exploitdb = df_exploitdb[['id', 'codes', 'description', 'platform', 'author', 'date_published']].copy()
        df_exploitdb = df_exploitdb.rename(columns={'id': 'exploit_id', 'codes': 'cve', 'description': 'exploit_description', 'platform': 'exploit_platform', 'author': 'exploit_author', 'date_published': 'exploit_date_published'})

        df_exploitdb['cve'] = df_exploitdb['cve'].astype(str).apply(lambda x: [c.strip().upper() for c in x.split(',') if c.strip().upper().startswith('CVE-')])
        df_exploitdb = df_exploitdb.explode('cve')
        df_exploitdb = df_exploitdb[df_exploitdb['cve'].notna() & (df_exploitdb['cve'] != '')]
        
        df_exploitdb['exploit_db_url'] = df_exploitdb['exploit_id'].apply(
            lambda x: f"https://www.exploit-db.com/exploits/{int(x)}" if pd.notna(x) else ''
        )

        if not shellcodes_df.empty:
            df_exploitdb = pd.concat([df_exploitdb, shellcodes_df], ignore_index=True)
            df_exploitdb = df_exploitdb.drop_duplicates(subset=['cve', 'exploit_id']) 

        _exploitdb_data = df_exploitdb
        loading_bar(current_overall_step + 1, total_overall_steps, overall_prefix, suffix=f"Exploit-DB carregado ({len(_exploitdb_data)} CVEs).", length=50)
        return _exploitdb_data

    except pd.errors.EmptyDataError as e:
        print(f"[\033[1;31m!\033[m] Erro no Exploit-DB CSV principal: {e}")
        import traceback; traceback.print_exc()
        return pd.DataFrame()
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao carregar/processar CSVs do Exploit-DB: {e}")
        import traceback; traceback.print_exc()
        return pd.DataFrame()

def delete_exploitdb_csv_files():
    if DELETE_EXPLOITDB_CSV_AFTER_USE:
        files_to_delete = [EXPLOITDB_EXPLOITS_LOCAL_CSV_PATH, EXPLOITDB_SHELLCODES_LOCAL_CSV_PATH]
        deleted_any = False
        for f_path in files_to_delete:
            if os.path.exists(f_path):
                try:
                    os.remove(f_path)
                    deleted_any = True
                except OSError as e:
                    print(f"[\033[1;31m!\033[m] Erro ao apagar arquivo '{f_path}': {e}")
                    import traceback; traceback.print_exc()
        if not deleted_any:
            print("[\033[1;34m[!]\033[m] Nenhum arquivo CSV do Exploit-DB para apagar ou já foram apagados.")
    else:
        print("[\033[1;34m[!]\033[m] Configuração para manter os arquivos CSV do Exploit-DB após o uso. Não serão apagados.")