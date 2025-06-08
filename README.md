# Conversor-de-XML-em-Danfe-

## 🔥 Processador de NF-e em Massa

**Aplicativo Python para processamento em massa de XMLs de NF-e gerando DANFEs em PDF**

Desenvolvido por **Thucosta** -

## ⚡ Características

- **Processamento em massa** de milhares de XMLs
- **Otimizado para velocidade** com cache de template
- **Apenas PDFs** - sem arquivos HTML desnecessários
- **Organização simples** - todos PDFs em uma pasta
- **Informações completas** - captura todos dados adicionais
- **Interface moderna** CustomTkinter
- **Progresso em tempo real** com estatísticas

## 🚀 Performance

**Otimizações implementadas:**
- ✅ Template carregado 1x (cache)
- ✅ Processador reutilizado
- ✅ Método otimizado para massa
- ✅ Log inteligente adaptativo
- ✅ **40-50% mais rápido** que versões anteriores

**Para 5000 XMLs:** ~1.5-2 horas (vs 2.5-3h antes)

## 📦 Instalação

```bash
pip install -r requirements.txt
```

## 🎯 Como usar

1. Execute o aplicativo:
```bash
python app_massa.py
```

2. **Selecione pasta com XMLs** de NF-e
3. **Configure pasta de saída** (padrão: `./massa_output`)
4. **Clique "Iniciar Processamento"**
5. **Acompanhe progresso** em tempo real

## 📁 Estrutura de Saída

```
massa_output/
├── XML_001.pdf
├── XML_002.pdf
├── XML_003.pdf
└── ...
```

Cada PDF é nomeado com base no arquivo XML original.

## 🛠️ Arquivos do Projeto

- `app_massa.py` - **Aplicativo principal**
- `processar_nfe.py` - Engine de processamento
- `nfe_vertical.html` - Template DANFE
- `requirements.txt` - Dependências

## 💡 Recursos Técnicos

- **CustomTkinter** - Interface moderna
- **WeasyPrint** - Geração de PDF
- **lxml** - Parse de XML
- **Threading** - Processamento em background
- **Cache inteligente** - Otimização de memória

## 📊 Dados Processados

**Informações capturadas:**
- ✅ Dados do emitente e destinatário
- ✅ Itens da nota fiscal
- ✅ Impostos (ICMS, IPI, PIS, COFINS)
- ✅ Duplicatas de cobrança
- ✅ **Informações adicionais completas** (fisco + complementares)
- ✅ Chave de acesso e protocolo

## 🎮 Interface

**Funcionalidades:**
- Seleção de pastas intuitiva
- Progresso visual em tempo real
- Log detalhado de mensagens
- Controle de parada
- Estatísticas finais

---

**Desenvolvido para processamento industrial de NF-es com foco em performance e confiabilidade.**