# Conversor de NF-e XML para DANFE HTML/PDF

Este projeto converte um arquivo XML de Nota Fiscal Eletrônica (NF-e) para um Documento Auxiliar da Nota Fiscal Eletrônica (DANFE) em formato HTML e PDF, utilizando um template HTML personalizável.

## Funcionalidades

*   Leitura de dados de um arquivo XML de NF-e.
*   Preenchimento de um template HTML (`nfe_vertical.html`) com os dados da NF-e.
*   Geração de um arquivo HTML (`danfe_gerada.html`) com a DANFE formatada.
*   Conversão do arquivo HTML gerado para um arquivo PDF (`danfe_gerada.pdf`) utilizando a biblioteca WeasyPrint.
*   Mapeamento flexível de campos do XML para o template HTML.
*   Formatação de valores monetários, datas e unidades de medida.

## Arquivos do Projeto

*   `processar_nfe.py`: Script principal em Python que realiza a leitura do XML, o preenchimento do template HTML e a conversão para PDF.
*   `nfe_vertical.html`: Template HTML base para a DANFE. Este arquivo pode ser personalizado para alterar o layout e a aparência da DANFE.
*   `26250512420164001048550010003065961741974330-nfe.xml`: Arquivo XML de exemplo de uma NF-e. Você pode substituí-lo pelo seu próprio arquivo XML. (O nome do arquivo XML a ser processado está definido no script `processar_nfe.py`).
*   `danfe_gerada.html` (Gerado): Arquivo HTML resultante após o processamento do XML e preenchimento do template.
*   `danfe_gerada.pdf` (Gerado): Arquivo PDF resultante da conversão do `danfe_gerada.html`.
*   `README.md`: Este arquivo.

## Pré-requisitos

1.  **Python:** Versão 3.7 ou superior. Você pode baixá-lo em [python.org](https://www.python.org/).
2.  **Biblioteca Python WeasyPrint:** Usada para a conversão de HTML para PDF.
3.  **Dependências de Sistema para WeasyPrint:** WeasyPrint depende de bibliotecas de sistema como Pango, Cairo e GDK-PixBuf. A instalação dessas dependências varia conforme o sistema operacional:
    *   **Windows:** Requer a instalação do GTK+. Siga as instruções em: [WeasyPrint on Windows](https://doc.weasyprint.org/stable/first_steps.html#windows).
    
    ## Configuração e Execução

1.  **Clone ou baixe este projeto** para o seu computador.
2.  **Instale a biblioteca WeasyPrint:**
    Abra um terminal ou prompt de comando e execute:
    ```bash
    pip install WeasyPrint
    ```
3.  **Instale as dependências de sistema do WeasyPrint** conforme descrito na seção "Pré-requisitos" e na documentação oficial do WeasyPrint. Este passo é crucial, especialmente no Windows.
4.  **Coloque o seu arquivo XML da NF-e** no mesmo diretório do script `processar_nfe.py`.
    *   Por padrão, o script está configurado para ler o arquivo `26250512420164001048550010003065961741974330-nfe.xml`.
    *   Se o seu arquivo XML tiver um nome diferente, você pode:
        *   Renomear o seu arquivo para `26250512420164001048550010003065961741974330-nfe.xml`, OU
        *   Editar a variável `xml_file_path` no início da função `main()` dentro do script `processar_nfe.py` para que aponte para o nome do seu arquivo.
5.  **Execute o script Python:**
    Navegue até o diretório do projeto no seu terminal/prompt de comando e execute:
    ```bash
    python processar_nfe.py
    ```
6.  **Verifique os arquivos gerados:**
    *   `danfe_gerada.html`: Abra este arquivo em um navegador para visualizar a DANFE em HTML.
    *   `danfe_gerada.pdf`: Abra este arquivo com um leitor de PDF para visualizar a DANFE em PDF.

## Personalização

*   **Layout da DANFE:** Para modificar a aparência da DANFE, edite o arquivo `nfe_vertical.html`. Você pode alterar o CSS, a estrutura das tabelas, adicionar ou remover campos (lembre-se de ajustar os placeholders e o script Python se adicionar novos campos que precisam ser preenchidos pelo XML).
*   **Mapeamento XML para HTML:** Se a estrutura do seu XML de NF-e for diferente ou se você precisar mapear outros campos, ajuste o dicionário `replacements` e as funções de formatação (como `format_items_html` e `format_duplicates_html`) no script `processar_nfe.py`.
*   **Nomes dos arquivos de saída:** Você pode alterar os nomes dos arquivos de saída (`output_html_path` e `output_pdf_path`) na função `main()` do script `processar_nfe.py`.

## Solução de Problemas

*   **Erro `OSError: cannot load library ...` ou similar (WeasyPrint):**
    Este erro geralmente indica que as dependências de sistema do WeasyPrint (GTK+, Pango, Cairo) não estão instaladas corretamente ou não estão acessíveis no PATH do sistema. Revise cuidadosamente as instruções de instalação do WeasyPrint para o seu sistema operacional. Certifique-se de reiniciar o terminal (ou o computador, se necessário no Windows) após instalar/atualizar o PATH.
*   **Campos em branco na DANFE:**
    Se alguns campos estiverem em branco, verifique:
    1.  Se o placeholder correspondente existe no `nfe_vertical.html`.
    2.  Se o caminho XPath para aquele campo no dicionário `replacements` em `processar_nfe.py` está correto para a estrutura do seu arquivo XML.
    3.  Se o campo realmente existe no seu arquivo XML de entrada. O script foi configurado para deixar campos não encontrados em branco.

---

ThTweaks ©