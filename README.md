# Estate Gover PPTX Engine

API para conectar o GPT da Estate Gover a uma automação externa de PowerPoint.

## Pastas

- `templates/`: colocar os modelos oficiais PPTX:
  - `estate-gover-capa-modelo-final.pptx`
  - `estate-gover-estrategica-modelo-final.pptx`
- `outputs/`: arquivos gerados.

## Rodar localmente

```bash
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 10000
```

## Endpoint

```text
POST /gerar-apresentacao-estate-gover
```

## Observação

Esta versão inicial duplica os modelos oficiais e devolve os links dos arquivos.
A próxima versão pode implementar substituição automática controlada de textos e imagens.
