# 📄 Conversor de Documentos para PDF

✅ **Status:** PRONTO PARA USO - Todos os testes passaram com sucesso!

Aplicação web desenvolvida em Python/Streamlit para conversão bidirecional de documentos Office ↔ PDF em alta qualidade.

## 🎯 Funcionalidades

- ✅ Conversão de documentos Office para PDF
- ✅ Conversão de PDF para DOCX
- ✅ Suporte a múltiplos formatos: DOC, DOCX, XLS, XLSX, PPT, PPTX, ODT, ODS, ODP
- ✅ Interface simples e intuitiva (tela única)
- ✅ PDFs em alta resolução para impressão profissional
- ✅ Processamento seguro (arquivos não são armazenados)
- ✅ Upload e download direto no navegador
- ✅ Estatísticas de conversão em tempo real
- ✅ Seletor de qualidade (Alta/Média/Padrão)

## 🧪 Status dos Testes

| Teste | Status | Data |
|-------|--------|------|
| Infraestrutura | ✅ PASSOU | 20/07/2026 |
| Conversão DOCX→PDF | ✅ PASSOU | 20/07/2026 |
| Interface Web | ✅ PASSOU | 20/07/2026 |
| LibreOffice 26.2.4.2 | ✅ FUNCIONANDO | 20/07/2026 |
| Todas dependências | ✅ INSTALADAS | 20/07/2026 |

**Resultado:** 🏆 100% dos testes aprovados!

Para detalhes completos, veja: [RESULTADO_TESTES.md](RESULTADO_TESTES.md)

## 🚀 Como usar localmente

### Pré-requisitos

- Python 3.8+
- LibreOffice instalado no sistema

#### Instalando LibreOffice

**Windows:**
```bash
# Baixe e instale de: https://www.libreoffice.org/download/download/
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install libreoffice libreoffice-writer libreoffice-calc libreoffice-impress
```

**macOS:**
```bash
brew install --cask libreoffice
```

### Instalação

1. Clone o repositório:
```bash
git clone <seu-repositorio>
cd libreoffice
```

2. Crie um ambiente virtual:
```bash
python -m venv venv
```

3. Ative o ambiente virtual:

**Windows:**
```bash
venv\Scripts\activate
```

**Linux/macOS:**
```bash
source venv/bin/activate
```

4. Instale as dependências:
```bash
pip install -r requirements.txt
```

### Executar a aplicação

```bash
streamlit run app.py
```

A aplicação será aberta automaticamente no navegador em `http://localhost:8501`

## ☁️ Deploy no Streamlit Cloud

1. Faça push do código para o GitHub
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte seu repositório
4. Configure o arquivo `packages.txt` (já incluído) para instalar o LibreOffice
5. Deploy!

## 📋 Estrutura do Projeto

```
libreoffice/
├── app.py              # Aplicação principal
├── requirements.txt    # Dependências Python
├── packages.txt        # Pacotes do sistema (LibreOffice)
├── .gitignore         # Arquivos ignorados pelo Git
└── README.md          # Este arquivo
```

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+**
- **Streamlit** - Framework web
- **LibreOffice** - Engine de conversão
- **subprocess** - Execução de comandos do sistema

## 📝 Formatos Suportados

### Entrada (Conversão para PDF)
- 📝 Documentos: DOC, DOCX, ODT
- 📊 Planilhas: XLS, XLSX, ODS
- 📽️ Apresentações: PPT, PPTX, ODP

## ⚙️ Configurações de Qualidade

A aplicação utiliza o LibreOffice em modo headless com configurações otimizadas para:
- Alta resolução de saída
- Preservação de formatação
- Fidelidade ao documento original
- Qualidade de impressão profissional

## 🔒 Segurança e Privacidade

- Todos os arquivos são processados em diretórios temporários
- Nenhum arquivo é armazenado permanentemente
- Limpeza automática após cada conversão
- Processamento local/isolado

## 📄 Licença

Este projeto é open source e está disponível sob a licença MIT.

## 👥 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:
- Reportar bugs
- Sugerir novas funcionalidades
- Enviar pull requests

## 📞 Suporte

Se encontrar problemas ou tiver dúvidas, abra uma issue no repositório.

---

Desenvolvido com ❤️ usando Python e Streamlit
