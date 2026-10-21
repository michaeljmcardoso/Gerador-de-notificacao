# Gerador de minuta de notificação

Aplicação Streamlit para preencher os dados do destinatário e do imóvel e gerar
uma minuta de notificação em DOCX, seguindo o modelo fornecido.

O formulário utiliza `data/processos_notificacao.db`, um cadastro derivado da
tabela `processos` que preserva os 445 registros e seus IDs, incluindo
duplicatas, e contém somente os dados necessários: número do processo,
comunidade e município. O seletor começa vazio e permite digitar para localizar
uma comunidade ou escolher uma opção. Ao selecionar um registro, o processo e o
município correspondentes são exibidos imediatamente. Comunidades com nomes
repetidos são identificadas também pelo município e, quando necessário, pelo
processo e pelo ID do registro.

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
