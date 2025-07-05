# exporters.py

import os
import pandas as pd
from datetime import datetime
import arrow 

from utils import garantir_colunas, formatar_aba, formatar_excel_output, clean_field, validar_credenciais, \
                    groups_to_str, convert_timestamp_to_datetime_str, prepare_df, is_group_empty, \
                    padronizar_excel_agents, loading_bar, inventario_software, \
                    load_exploitdb_data, delete_exploitdb_csv_files, TenableIO 

from config import STATUS_OFFLINE, TABLE_STYLE, DELETE_EXPLOITDB_CSV_AFTER_USE 


########## assets export ###################################
def exportar_assets(tio: TenableIO, output_folder: str, cliente: str, assets_data: list = None):

    try:
        total_passos = 7
        passo = 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Início', length=50)

        assets = assets_data if assets_data is not None else list(tio.exports.assets())
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Coletando/Usando dados', length=50)

        df = pd.json_normalize(assets, sep='.')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Normalizando dados', length=50)
        cols_to_clean = ['agent_names', 'fqdns', 'ipv4s', 'ipv6s', 'mac_addresses',
                         'operating_systems', 'hostnames', 'netbios_names', 'sources', 'system_types']
        
        for col in cols_to_clean:
            if col in df.columns:
                df[col] = df[col].apply(clean_field)

        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Limpando dados selecionados', length=50)

        colunas_desejadas = [
            'ratings.acr.score', 'acr_score', 'exposure_score', 'agent_names', 'fqdns',
            'ipv4s', 'ipv6s', 'mac_addresses', 'operating_systems', 'first_seen', 'has_agent',
            'hostnames', 'id', 'last_authenticated_scan_date', 'last_licensed_scan_date',
            'last_seen', 'name', 'netbios_names', 'bios_uuid', 'servicenow_sysid',
            'sources', 'system_types', 'agent_uuid'
        ]
        df = garantir_colunas(df, colunas_desejadas)
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Garantindo colunas', length=50)

        df_final = df.reindex(columns=colunas_desejadas)
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Reorganizando colunas', length=50)

        output_file = os.path.join(output_folder, f'Tenable_assets_{cliente}.xlsx')
        formatar_excel_output(df_final, output_file, 'Assets')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Salvando arquivo', length=50)

        print(f"\n[\033[1;32m✓\033[m] Exportação de assets concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar assets: {e}")
        traceback.print_exc()

############## agents export ###################################
def exportar_agents(tio: TenableIO, output_folder: str, cliente: str, filtro: str = 'offline', agents_data: list = None):

    try:
        is_compare = filtro == 'compare'
        total_passos = 11 if is_compare else 6 
        passo = 1

        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Início', length=50)

        agents = agents_data if agents_data is not None else list(tio.agents.list())
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Listando/Usando agents', length=50)

        df = pd.json_normalize(agents, sep='.')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Normalizando dados', length=50)

        if is_compare:
            offline_df = groups_to_str(df[df['status'].str.lower() == STATUS_OFFLINE])
            passo += 1
            loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Filtrando Offline', length=50)

            nogroup_df = groups_to_str(df[df['groups'].apply(is_group_empty)])
            passo += 1
            loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Filtrando Sem Grupo', length=50)

            intersection_df = groups_to_str(df[
                (df['status'].str.lower() == STATUS_OFFLINE) & (df['groups'].apply(is_group_empty))
            ])
            passo += 1
            loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Unindo filtros', length=50)

            union_df = pd.concat([offline_df, nogroup_df]).drop_duplicates()
            todos_df = groups_to_str(df)

            output_file = os.path.join(output_folder, f'Tenable_agents_compare_{cliente}.xlsx')
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                dfs = {
                    'Offline': offline_df,
                    'SemGrupo': nogroup_df,
                    'Offline_ou_SemGrupo': union_df,
                    'Offline_e_SemGrupo': intersection_df,
                    'Todos': todos_df
                }
                aba_count = 0
                for aba, df_sub in dfs.items():
                    aba_count += 1
                    df_final = prepare_df(df_sub)
                    df_final.to_excel(writer, sheet_name=aba, index=False)
                    formatar_aba(writer, df_final, aba)
                    
                    current_suffix = f'Formatando aba {aba}'
                    if aba_count == len(dfs): 
                        current_suffix = 'Concluído' 
                        
                    loading_bar(passo + aba_count, total_passos, prefix='Exportando Agents:', suffix=current_suffix, length=50)

            print(f"\n[\033[1;32m✓\033[m] Exportação concluída: {output_file}")

        else: 
            if filtro == 'offline':
                filtrados = df[df['status'].str.lower() == STATUS_OFFLINE].copy()
                filtro_desc = 'Offline'
            elif filtro == 'nogroup':
                filtrados = df[df['groups'].apply(is_group_empty)].copy()
                filtro_desc = 'Sem Grupo'
            else: 
                filtrados = df.copy()
                filtro_desc = 'Todos'

            passo += 1
            loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix=f'Filtrando: {filtro_desc}', length=50)

            filtrados = groups_to_str(filtrados)
            df_final = prepare_df(filtrados)
            output_file = os.path.join(output_folder, f'Tenable_agents_{filtro}_{cliente}.xlsx')

            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Agents', index=False)
                formatar_aba(writer, df_final, 'Agents')
                passo += 1
                loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Concluído', length=50) 
            
            print(f"\n[\033[1;32m✓\033[m] Exportação concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar agents: {e}")
        traceback.print_exc()

###### vuln export ###################################

def exportar_vulnerabilidades(tio: TenableIO, output_folder: str, cliente: str, filtro_severidade: list = None, last_found_days: int = None, include_exploitdb: bool = False):
    try:
        total_passos = 8 
        
        passo = 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Início', length=50)

        filters = {}
        if filtro_severidade:
            filters['severity'] = filtro_severidade
        if last_found_days: 
            filters['last_found'] = int(arrow.now().shift(days=-last_found_days).timestamp())

        vulnerabilities = list(tio.exports.vulns(**filters))
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Coletando dados da API', length=50)

        if not vulnerabilities:
            print("\n[\033[1;33m!\033[m] Nenhuma vulnerabilidade encontrada com os filtros especificados.")
            return

        df_vuln = pd.json_normalize(vulnerabilities, sep='.')
        
        df_vuln = garantir_colunas(df_vuln, ['asset.id'])


        passo += 1 
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Normalizando dados de vulnerabilidades', length=50)

        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Coletando dados completos de assets', length=50)
        all_assets_full = list(tio.exports.assets()) 
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix=f'Processando {len(all_assets_full)} assets', length=50)
        df_assets = pd.json_normalize(all_assets_full, sep='.') 
        
        df_assets = garantir_colunas(df_assets, ['id'])


        asset_name_fields_for_merge = [
            'hostnames',     
            'name',          
            'fqdns',         
            'ipv4s',         
            'netbios_names', 
            'uuid'           
        ]
        
        df_assets['Master Asset Name'] = ''
        for field in asset_name_fields_for_merge:
            if field in df_assets.columns:
                df_assets['Master Asset Name'] = df_assets.apply(
                    lambda row: clean_field(row[field]) if row['Master Asset Name'] == '' else row['Master Asset Name'],
                    axis=1
                )
            df_assets['Master Asset Name'] = df_assets['Master Asset Name'].apply(clean_field)
        
        df_assets['Master Asset Name'] = df_assets['Master Asset Name'].apply(lambda x: 'N/A' if x == '' else x)
        
        df_assets_for_merge = df_assets[['id', 'Master Asset Name']].copy()
        
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Mesclando dados de assets e vulnerabilidades', length=50)
        df = pd.merge(df_vuln, df_assets_for_merge, left_on='asset.id', right_on='id', how='left', suffixes=('_vuln', '_asset'))
        
        if 'id_asset' in df.columns:
            df = df.drop(columns=['id_asset'])
        
        df.rename(columns={'Master Asset Name': 'Asset Name'}, inplace=True)

        if 'Asset Name' not in df.columns:
             df['Asset Name'] = 'N/A'

        df['cve_id'] = None
        if 'cve' in df.columns:
            df['cve_id'] = df['cve'].apply(lambda x: x[0] if isinstance(x, list) and x else None)
        elif 'plugin.cve' in df.columns:
            df['cve_id'] = df['plugin.cve'].apply(lambda x: x[0] if isinstance(x, list) and x else None)
        
        df_exploitdb = pd.DataFrame()
        if include_exploitdb:
            total_passos += 2 
            df_exploitdb = load_exploitdb_data(passo, total_passos, 'Exportando Vulnerabilidades:')
            passo += 1 
            
            if not df_exploitdb.empty and 'cve' in df_exploitdb.columns and 'cve_id' in df.columns and df['cve_id'].notna().any():
                loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Correlacionando com Exploit-DB', length=50)
                df = pd.merge(df, df_exploitdb, left_on='cve_id', right_on='cve', how='left', suffixes=('_tenable', '_exploitdb'))
                
                df['Exploit Disponível (Exploit-DB)'] = df['exploit_id'].notna().map({True: 'Sim', False: 'Não'})
                df['Exploit-DB Link'] = df['exploit_db_url'].apply(lambda x: str(x)[:255] if pd.notna(x) else '') 
            else:
                print("[\033[1;33m!\033[m] Aviso: Exploit-DB selecionado, mas não foi possível correlacionar (base de dados vazia ou sem CVEs). Continuando sem dados do Exploit-DB.")
            passo += 1

        cols_to_clean = [
            'asset.agent_names', 'asset.fqdns', 'asset.ipv4s', 'asset.ipv6s', 
            'asset.mac_addresses', 'asset.operating_systems', 'asset.netbios_names', 'asset.sources', 'asset.system_types',
            'plugin.family', 'plugin.name', 'severity', 'state', 'output'
        ]
        for field in ['asset.name', 'asset.fqdns', 'asset.ipv4s', 'asset.hostnames', 'asset.netbios_names', 'asset.uuid']:
            if field in cols_to_clean:
                cols_to_clean.remove(field)

        for col in cols_to_clean:
            if col in df.columns:
                df[col] = df[col].apply(clean_field)
        
        for date_col in ['first_found', 'last_found']:
            if date_col in df.columns:
                df[date_col] = df[date_col].apply(convert_timestamp_to_datetime_str)

        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Limpando e formatando dados', length=50)
        
        colunas_desejadas = [
            'plugin.name', 'plugin.id', 'severity', 
            'asset.hostname', 'asset.ipv4',  
            'first_found', 'last_found', 'output', 'state', 'asset.id', 'plugin.family', 'cve_id'
        ]
        
        if include_exploitdb and not df_exploitdb.empty and 'Exploit Disponível (Exploit-DB)' in df.columns:
            colunas_exploitdb_a_adicionar = [
                'Exploit Disponível (Exploit-DB)',
                'Exploit-DB Link',
                'exploit_id',      
            ]
            colunas_desejadas.extend(colunas_exploitdb_a_adicionar)
            
        else: 
            for col_name_raw in ['Exploit Disponível (Exploit-DB)', 'Exploit-DB Link', 'exploit_id']:
                col_name_display = col_name_raw 
                if col_name_display not in df.columns:
                    df[col_name_display] = ''


        df = garantir_colunas(df, colunas_desejadas)
        df_final = df.reindex(columns=colunas_desejadas)
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Reorganizando colunas', length=50)

        output_file = os.path.join(output_folder, f'Tenable_vulnerabilities_{cliente}.xlsx')
        formatar_excel_output(df_final, output_file, sheet_name='Vulnerabilities')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Vulnerabilidades:', suffix='Salvando arquivo', length=50)

        print(f"\n[\033[1;32m✓\033[m] Exportação de vulnerabilidades concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar vulnerabilidades: {e}")
        traceback.print_exc()
    finally:
        if include_exploitdb: 
            delete_exploitdb_csv_files()