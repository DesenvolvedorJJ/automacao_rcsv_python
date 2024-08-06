import pandas as pd
import glob
import os

def process_csv_file(file_path):
    # Lê o arquivo CSV
    df = pd.read_csv(file_path)
    
    # Trata IP vazio como 'DESCONHECIDO' para o agrupamento
    df['Endereço de IP'] = df['Endereço de IP'].fillna('DESCONHECIDO')
    
    # Mapear os modelos para um grupo consolidado
    model_mapping = {
        'M337x 387X 407X': 'M337x 387X 407X - Sl-M4070fr',
        'M337x387X407X': 'M337x 387X 407X - Sl-M4070fr',
        'M337x': 'M337x 387X 407X - Sl-M4070fr',
        '387X': 'M337x 387X 407X - Sl-M4070fr',
        '407X': 'M337x 387X 407X - Sl-M4070fr',
        'Sl-M4070fr': 'M337x 387X 407X - Sl-M4070fr'
    }
    
    df['Modelo Consolidado'] = df['Modelo'].map(model_mapping).fillna(df['Modelo'])
    
    # Agrupa os dados por Fabricante e Modelo Consolidado
    grouped = df.groupby(['Fabricante', 'Modelo Consolidado', 'Endereço de IP'])
    
    # Cria uma lista para armazenar os resultados
    results = []
    
    for (fabricante, modelo, ip), group in grouped:
        total_producao = group['Producao'].sum()
        qtde_impressoras = len(group['Endereço de IP'].unique())
        results.append({
            'Fabricante': fabricante,
            'Modelo': modelo,
            'Qtde.Impressão': total_producao,
            'Qtde.Impressoras': qtde_impressoras,
            'Média': total_producao / qtde_impressoras if qtde_impressoras > 0 else 0
        })
    
    results_df = pd.DataFrame(results)
    
    # Agrupa novamente por Fabricante e Modelo Consolidado para obter os totais
    final_results = results_df.groupby(['Fabricante', 'Modelo']).agg({
        'Qtde.Impressão': 'sum',
        'Qtde.Impressoras': 'sum'
    }).reset_index()
    
    final_results['Média'] = final_results['Qtde.Impressão'] / final_results['Qtde.Impressoras']
    
    # Calcula totais gerais
    total_general = final_results[['Qtde.Impressão', 'Qtde.Impressoras', 'Média']].sum()
    total_general['Fabricante'] = 'TOTAL'
    total_general['Modelo'] = ''
    
    # Converte o total_general em um DataFrame para concatenação
    total_general_df = pd.DataFrame([total_general])
    
    # Adiciona a linha de total ao DataFrame final
    final_results = pd.concat([final_results, total_general_df], ignore_index=True)
    
    return final_results

def save_to_excel(dataframe, output_file):
    dataframe.to_excel(output_file, index=False)

def main():
    # Especifique o caminho dos arquivos CSV
    csv_files = glob.glob('mes/*.csv')
    
    for file_path in csv_files:
        # Processa o arquivo CSV
        results = process_csv_file(file_path)
        
        # Cria o nome do arquivo Excel com base no nome do arquivo CSV
        base_name = os.path.basename(file_path)
        excel_file = os.path.splitext(base_name)[0] + '_media_por_model.xlsx'
        
        # Salva os resultados em um arquivo Excel
        save_to_excel(results, excel_file)
        print(f'Arquivo gerado: {excel_file}')

if __name__ == "__main__":
    main()
