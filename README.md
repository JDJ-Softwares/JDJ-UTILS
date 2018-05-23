# JDJ-UTILS

## Objetivo

	Consiste em uma série de classes para manipulação de arquivos tendo em vista a facilitação de importação e exportação destes.
	É uma biblioteca experimental em desenvolvimento, ainda não têm uma documentação específica.
	
## Builds

### 0.0.1-Alpha

#### Ferramentas

	*PDFUnion - Junção de arquivos em .pdf
	*Excel-Util - importador e exportador de arquivos em .xls
	
### Implementação
	
	Crie um contrutor para receber um Map<K, V>:
	
	public Pessoa(Map<String, String> map) {
		super();
		this.nome = map.get("nome");
		this.sobrenome = map.get("sobrenome");
	}
	
	Certifique-se de que o Arquivo esteja com suas colunas com o mesmo nome das variáveis do DTO.