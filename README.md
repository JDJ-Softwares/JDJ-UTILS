# JDJ-UTILS

## Objetivo

	Consiste em uma s�rie de classes para manipula��o de arquivos tendo em vista a facilita��o de importa��o e exporta��o destes.
	� uma biblioteca experimental em desenvolvimento, ainda n�o t�m uma documenta��o espec�fica.
	
## Builds

### 0.0.1-Alpha

#### Ferramentas

	*PDFUnion - Jun��o de arquivos em .pdf
	*Excel-Util - importador e exportador de arquivos em .xls
	
### Implementa��o
	
	Crie um contrutor para receber um Map<K, V>:
	
	public Pessoa(Map<String, String> map) {
		super();
		this.nome = map.get("nome");
		this.sobrenome = map.get("sobrenome");
	}
	
	Certifique-se de que o Arquivo esteja com suas colunas com o mesmo nome das vari�veis do DTO.