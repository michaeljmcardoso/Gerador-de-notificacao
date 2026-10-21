# Gerador de minuta de notificação

Aplicação Streamlit para preencher os dados do destinatário e do imóvel e gerar
uma minuta de notificação em DOCX, seguindo o modelo fornecido.

## Executar

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

O código do imóvel no cabeçalho e o código informado no segundo parágrafo são
campos independentes, pois podem ser diferentes, como no modelo. O nome do
imóvel e a matrícula preenchidos no formulário são utilizados nos dois locais.
Os dois campos de código são opcionais; quando não preenchidos, são omitidos da
minuta.
