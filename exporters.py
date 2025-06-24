import os
import pandas as pd
from datetime import datetime

STATUS_ONLINE = 'on'
STATUS_OFFLINE = 'off'
TABLE_STYLE = 'Table Style Medium 2'

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
        print(f"[\033[1m1/4\033[0m] Salvando planilha Excel: {output_file}")
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=aba, index=False)
            formatar_aba(writer, df, aba)
        print(f"[\033[1m3/4\033[0m] Ajuste das colunas concluído.")
    except Exception as e:
        fallback = output_file.replace('.xlsx', '.csv')
        df.to_csv(fallback, index=False)
        print(f"[\033[1;31m!\033[m] Erro ao salvar Excel ({e}). CSV criado: {fallback}")
    print(f"[\033[1m4/4\033[0m] Exportação finalizada.")

def clean_field(val):
    """Limpa campos do tipo lista ou dict para strings amigáveis no DataFrame."""
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

def exportar_assets(tio, output_folder, cliente):
    try:
        print("[\033[1m1/4\033[0m] Iniciando exportação de assets...")
        assets = list(tio.exports.assets())
        print(f"[✓] {len(assets)} ativos coletados.")

        df = pd.json_normalize(assets, sep='.')
        df = df.applymap(clean_field)

        colunas_desejadas = [
            'ratings.acr.score', 'acr_score', 'exposure_score', 'agent_names', 'fqdns',
            'ipv4s', 'ipv6s', 'mac_addresses', 'operating_systems', 'first_seen', 'has_agent',
            'hostnames', 'id', 'last_authenticated_scan_date', 'last_licensed_scan_date',
            'last_seen', 'name', 'netbios_names', 'bios_uuid', 'servicenow_sysid',
            'sources', 'system_types', 'agent_uuid'
        ]

        df = garantir_colunas(df, colunas_desejadas)
        df_final = df.reindex(columns=colunas_desejadas)

        output_file = os.path.join(output_folder, f'Tenable_assets_{cliente}.xlsx')
        formatar_excel(df_final, output_file, 'Assets')
        print(f"[\033[1;32m✓\033[m] Exportação de assets concluída: {output_file}")

    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao exportar assets: {e}")

def exportar_agents(tio, output_folder, cliente, filtro='offline'):
    try:
        print("[\033[1m1/4\033[0m] Iniciando exportação de agents...")
        agents = list(tio.agents.list())
        df = pd.json_normalize(agents, sep='.')

        def is_group_empty(x):
            if isinstance(x, list):
                return len(x) == 0
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return True
            if isinstance(x, str) and x.strip() == '':
                return True
            return False

        def groups_to_str(dff):
            dff = dff.copy()
            dff['groups'] = dff['groups'].apply(
                lambda v: ', '.join([g.get('name', '') for g in v]) if isinstance(v, list) else '')
            return dff

        def ts(t):
            try:
                return datetime.fromtimestamp(int(t)).strftime('%Y-%m-%d %H:%M:%S') if t else ''
            except:
                return ''

        def prepare_df(dff):
            dff = dff.copy()
            if 'core_version' in dff.columns:
                dff['version_agent'] = dff['core_version'].fillna('N/A').replace('', 'N/A')
            else:
                dff['version_agent'] = 'N/A'

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

        print(f"[\033[1m2/4\033[0m] Ajuste das colunas concluído.")

        if filtro == 'compare':
            filtro_str = 'compare'

            offline_df = groups_to_str(df[df['status'].str.lower() == STATUS_OFFLINE])
            nogroup_df = groups_to_str(df[df['groups'].apply(is_group_empty)])
            intersection_df = groups_to_str(df[
                (df['status'].str.lower() == STATUS_OFFLINE) & (df['groups'].apply(is_group_empty))
            ])
            union_df = pd.concat([offline_df, nogroup_df]).drop_duplicates()
            todos_df = groups_to_str(df)

            print(f"\n[\033[1;32m✓\033[m] Total agentes offline: \033[1m{len(offline_df)}\033[0m")
            print(f"[\033[1;32m✓\033[m] Total agentes sem grupo: \033[1m{len(nogroup_df)}\033[0m")
            print(f"[\033[1;32m✓\033[m] Total agentes offline ou sem grupo (união): \033[1m{len(union_df)}\033[0m")
            print(f"[\033[1;32m✓\033[m] Total agentes offline e sem grupo (interseção): \033[1m{len(intersection_df)}\033[0m")

            output_file = os.path.join(output_folder, f'Tenable_agents_{filtro_str}_{cliente}.xlsx')
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                dfs = {
                    'Offline': offline_df,
                    'SemGrupo': nogroup_df,
                    'Offline_ou_SemGrupo': union_df,
                    'Offline_e_SemGrupo': intersection_df,
                    'Todos': todos_df
                }
                print(f"[\033[1m3/4\033[0m] Aplicando estilo azul na tabela")
                for aba, df_sub in dfs.items():
                    df_final = prepare_df(df_sub)
                    df_final.to_excel(writer, sheet_name=aba, index=False)
                    formatar_aba(writer, df_final, aba)

            print(f"[\033[1m4/4\033[0m] Finalizando exportação comparativa de agents...")
            print(f"\n[\033[1;32m✓\033[m] Exportação comparativa concluída: {output_file}")

        else:
            if filtro == 'offline':
                filtrados = df[df['status'].str.lower() == STATUS_OFFLINE].copy()
                filtro_str = 'offline'
                filtrados = groups_to_str(filtrados)
            elif filtro == 'nogroup':
                filtrados = df[df['groups'].apply(is_group_empty)].copy()
                filtro_str = 'nogroup'
                filtrados = groups_to_str(filtrados)
            else:
                filtrados = groups_to_str(df.copy())
                filtro_str = 'all'

            df_final = prepare_df(filtrados)
            output_file = os.path.join(output_folder, f'Tenable_agents_{filtro_str}_{cliente}.xlsx')
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, sheet_name='Agents', index=False)
                print(f"[\033[1m3/4\033[0m] Aplicando estilo azul na tabela")
                formatar_aba(writer, df_final, 'Agents')

            print(f"[\033[1m4/4\033[0m] Finalizando exportação de agents...")
            print(f"\n[\033[1;32m✓\033[m] Exportação concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar agents: {e}")
        traceback.print_exc()

def validar_credenciais(tio):
    try:
        session = tio.session.details()
        print(f"[\033[1;32m✓\033[m] Conectado como: \033[1m{session['username']}\033[0m")
        return True
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Falha na autenticação: {e}")
        return False
