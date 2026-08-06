# 📄 Conversor de Formatos Antigos do Office

Aplicação web em Python/Streamlit que converte **formatos antigos do Office para os formatos modernos**, usando o LibreOffice em modo headless.

## 🎯 O que faz

| Entrada (antigo) | Saída (moderno) |
| --- | --- |
| `.doc` · `.odt` · `.rtf` | **`.docx`** (Word) |
| `.xls` · `.ods` | **`.xlsx`** (Excel) |
| `.ppt` · `.odp` | **`.pptx`** (PowerPoint) |

- Interface de tela única (upload → converter → baixar)
- Processamento em diretórios temporários — nenhum arquivo é armazenado

## 🚀 Rodar localmente

Pré-requisitos: Python 3.8+ e LibreOffice instalado.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abre em `http://localhost:8501`.

## ☁️ Deploy no Streamlit Cloud

1. Faça push para o GitHub
2. Em [share.streamlit.io](https://share.streamlit.io), conecte o repositório
3. O `packages.txt` (incluído) instala o LibreOffice automaticamente
4. Deploy

## 📋 Estrutura

```
doc-para-docx/
├── app.py            # Aplicação principal
├── requirements.txt  # Dependência Python (streamlit)
├── packages.txt      # Pacotes do sistema (LibreOffice)
└── README.md
```

## 🛠️ Tecnologias

- **Streamlit** — interface web
- **LibreOffice** (headless) — motor de conversão
- **subprocess** — execução do LibreOffice

## 🔒 Privacidade

Os arquivos são processados em diretórios temporários e removidos após a conversão. Nada é armazenado permanentemente.

---

Desenvolvido com ❤️ usando Python e Streamlit.
