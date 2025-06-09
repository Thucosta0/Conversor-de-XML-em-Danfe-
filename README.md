# Conversor NFe XML para PDF DANFE 📄

Aplicação Python moderna para conversão de arquivos XML de Notas Fiscais Eletrônicas (NFe) em PDFs no formato DANFE, com funcionalidades adicionais de renomeação em massa.

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Status](https://img.shields.io/badge/Status-Ativo-brightgreen?style=flat-square)

## 🚀 Funcionalidades

### ✨ Conversor XML → PDF
- **Conversão em massa** de arquivos XML NFe para PDF DANFE
- **Template profissional** com layout empresarial
- **Mapeamento completo** de todos os campos da NFe
- **Processamento otimizado** para grandes volumes
- **Barra de progresso** com estatísticas em tempo real
- **Log detalhado** das operações
- **Relatório Excel** automático com chave de acesso, número da NF e status de conversão

### 🔄 Renomeador Inteligente
- **Renomeação em massa** de arquivos XML e PDF
- **Interface tabular** intuitiva com TreeView
- **Validação automática** de chaves NFe (44 dígitos)
- **Filtros avançados** por texto e status
- **Processamento assíncrono** sem travamento da interface
- **Sistema de logs** com salvamento

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **CustomTkinter** - Interface moderna
- **lxml** - Processamento XML
- **weasyprint** - Geração de PDF
- **pandas** - Manipulação de dados
- **threading** - Processamento assíncrono

## 📦 Instalação

### Pré-requisitos

Certifique-se de ter o Python 3.8 ou superior instalado no sistema.

### Dependências

```bash
pip install -r requirements.txt
```

Ou instale manualmente:

```bash
pip install customtkinter lxml weasyprint pandas
```

## 🎯 Como Usar

### Executar a Aplicação

```bash
python app_massa.py
```

### Conversor XML → PDF

1. Clique em **"Conversor XML→PDF"**
2. Selecione a pasta contendo os arquivos XML
3. Defina a pasta de destino para os PDFs
4. Clique em **"Iniciar Conversão"**
5. Acompanhe o progresso no log em tempo real

### Renomeador XML/PDF

1. Clique em **"Renomeador XML/PDF"**
2. Selecione a pasta com arquivos XML/PDF
3. Clique em **"Adicionar Dados"**
4. Insira as chaves NFe e nomes desejados
5. Execute a validação e renomeação

## 📁 Estrutura do Projeto

```
Conversor-de-XML-em-Danfe-/
├── app_massa.py          # Aplicação principal
├── nfe_vertical.html     # Template HTML para DANFE
├── requirements.txt      # Dependências do projeto
└── README.md            # Documentação
```

## 🎨 Interface

A aplicação possui interface moderna construída com CustomTkinter, oferecendo:

- **Design responsivo** com cards organizados
- **Navegação por abas** entre funcionalidades
- **Feedback visual** em tempo real
- **Confirmações de segurança** para operações críticas
- **Filtros e busca** em tempo real

## 📋 Template DANFE

O template HTML inclui mapeamento completo para:

- **Dados da empresa emissora** (9 campos)
- **Informações da NFe** (8 campos)
- **Dados do destinatário** (9 campos)
- **Cálculos de impostos** (12 campos)
- **Informações de transporte** (16 campos)
- **Tabela de produtos** (13 colunas)
- **Duplicatas e parcelas**
- **Informações adicionais**

## 📊 Relatório Excel Profissional

Após cada conversão em massa, a aplicação gera automaticamente um relatório Excel completo com:

### 📋 Aba "Relatório Conversão"
- **Chave de Acesso** - Chave de 44 dígitos da NFe
- **Nota Fiscal** - Número da nota fiscal
- **Sucesso de Conversão** - Status "Sim" ou "Não"
- **Arquivo XML** - Nome do arquivo processado
- **Data/Hora Processamento** - Timestamp da operação
- **Pasta Origem** - Caminho do arquivo original
- **Tamanho Arquivo (KB)** - Tamanho do arquivo XML
- **Erro Detalhado** - Mensagem específica de erro (quando aplicável)

### 📈 Aba "Estatísticas"
- **Total de Arquivos** processados
- **Conversões Bem-sucedidas** e **com Erro**
- **Taxa de Sucesso** em porcentagem
- **Tamanho Total Processado** em MB
- **Tempo Total** e **Velocidade Média** de processamento
- **Data/Hora** de início e fim da operação

### 🎨 Formatação Avançada
- **Cores condicionais**: Verde para sucessos, vermelho para erros
- **Cabeçalhos estilizados** com cores profissionais
- **Colunas auto-ajustadas** para melhor visualização
- **Bordas e alinhamento** profissional

### 🚀 Abertura Automática
- **Pergunta automática** se deseja abrir o Excel após conclusão
- **Compatibilidade multiplataforma** (Windows, macOS, Linux)
- **Arquivo salvo** na pasta de destino dos PDFs

Nome do arquivo: `Relatorio_Conversao_NFe_YYYYMMDD_HHMMSS.xlsx`

## 🔧 Configuração Avançada

### Template Personalizado

O arquivo `nfe_vertical.html` pode ser customizado para atender necessidades específicas. O sistema mapeia automaticamente os campos XML para os placeholders `[campo]` no template.

### Logs

Os logs são salvos automaticamente e incluem:
- Timestamp das operações
- Status de processamento
- Erros e exceções
- Estatísticas de performance

## 🤝 Contribuição

Contribuições são bem-vindas! Para contribuir:

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📝 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 👨‍💻 Desenvolvedor

**Thucosta**

- GitHub: [@thucosta](https://github.com/thucosta)
- LinkedIn: [Arthur Costa](https://linkedin.com/in/thucosta)


## 📊 Status do Projeto

- ✅ Conversão XML → PDF funcional
- ✅ Interface moderna implementada
- ✅ Renomeador em massa funcional
- ✅ Sistema de logs implementado
- 🔄 Melhorias contínuas de performance
- 📋 Documentação em andamento

---

⭐ Se este projeto foi útil para você, não esqueça de dar uma estrela!
