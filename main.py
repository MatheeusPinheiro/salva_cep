import http.client
import json
import pandas as pd


def obter_endereco_cep(cep):
	"""
	Esta função cosulta a API do ViaCEP para obter informações de endereço com base em um cep.

	Parametros:
	cep (str): O CEP para o qual o endereço deve ser consultado.

	Retorna:
	dict ou str: Retorna um dicionário contedo as informações do endereço ou uma mensagem de erro se o CEP não for encontrado
	"""
	conexao = http.client.HTTPSConnection('viacep.com.br')

	conexao.request("GET", f"/ws/{cep}/json/")

	resposta = conexao.getresponse()
	
	if resposta.status != 200:
		conexao.close()
		return None

	dados = resposta.read()
	
	endereco = json.loads(dados.decode('utf-8'))
	
	conexao.close()
	
	return endereco if 'erro' not in endereco else None


def salvar_endereco_excel(endereco, nome_arquivo='endereco.xlsx'):
    if 'erro' not in endereco:
        df = pd.DataFrame([endereco])

        df.to_excel(nome_arquivo, index=False)

        print(f'Dados salvos com sucesso no arquivo {nome_arquivo}')
    else:
        print('Não foi possível salvar os dados: CEP não encontrado.')


def main():
	try:
		
		planilha_ceps = pd.read_excel('CEP.xlsx', sheet_name='CEP')
		ceps = planilha_ceps['CEP'].dropna()
  
		resultados = pd.DataFrame(columns=['CEP','Logradouro', 'Bairro', 'Localidade', 'UF'])
  
		for cep in ceps:
			endereco = obter_endereco_cep(str(cep).replace('-', ''))

			if endereco:
				nova_linha =  pd.DataFrame([{
					"CEP": cep,
					"Logradouro":endereco.get('logradouro', ''),
					"Bairro": endereco.get('bairro', ''),
					"Localidade":endereco.get('localidade', ''),
					"UF":endereco.get('uf', ''),
				}])
				resultados = pd.concat([resultados, nova_linha], ignore_index=True)
    
		with pd.ExcelWriter('CEP.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
			resultados.to_excel(writer, sheet_name='Dados', index=False)

		print("Endereços salvos na aba de 'Dados'")
  
	except Exception as error:
		print('Erro: ', error)
    
''

if __name__ == '__main__':
    main()