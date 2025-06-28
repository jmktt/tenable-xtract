import os
import pandas as pd
from datetime import datetime
from utils import *

def exportar_assets(tio, output_folder, cliente):
    try:
        total_passos = 7
        passo = 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Início', length=50)

        assets = list(tio.exports.assets())
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Coletando dados', length=50)

        df = pd.json_normalize(assets, sep='.')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Normalizando dados', length=50)

        df = df.apply(lambda col: col.map(clean_field))
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Limpando dados', length=50)

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
        formatar_excel(df_final, output_file, 'Assets')
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Assets:', suffix='Salvando arquivo', length=50)

        print(f"\n[\033[1;32m✓\033[m] Exportação de assets concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar assets: {e}")
        traceback.print_exc()


def exportar_agents(tio, output_folder, cliente, filtro='offline'):
    try:
        is_compare = filtro == 'compare'
        total_passos = 12 if is_compare else 6
        passo = 1

        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Início', length=50)

        agents = list(tio.agents.list())
        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Listando agents', length=50)

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
                for aba, df_sub in dfs.items():
                    df_final = prepare_df(df_sub)
                    df_final.to_excel(writer, sheet_name=aba, index=False)
                    formatar_aba(writer, df_final, aba)
                    passo += 1
                    loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix=f'Formatando aba {aba}', length=50)

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
                loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Formatando aba Agents', length=50)

        passo += 1
        loading_bar(passo, total_passos, prefix='Exportando Agents:', suffix='Concluído                        ', length=50)
        print(f"\n[\033[1;32m✓\033[m] Exportação concluída: {output_file}")

    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar agents: {e}")
        traceback.print_exc()
