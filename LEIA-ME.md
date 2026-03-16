# Prévia de Sustentações Orais — Gerador

Ferramenta offline para geração automática da lista de sustentações orais em formato Word (.docx).

## Como usar

1. Abra o arquivo `index.html` no navegador (Chrome recomendado)
2. Preencha os campos da sessão (turma, sala, data)
3. Selecione o arquivo Excel (.xlsx) exportado do Extrator
4. Clique em "Gerar Prévia (.docx)"
5. O arquivo Word será baixado automaticamente

## Estrutura da planilha esperada

A planilha deve conter as seguintes colunas (mesma estrutura do Sustentação Oral — Extrator):

| Coluna | Campo                        |
|--------|------------------------------|
| A      | Classe Judicial              |
| B      | Sigiloso                     |
| C      | Observações (idoso, etc.)    |
| D      | Nº Processo (CNJ)           |
| E      | Data Sessão                  |
| F      | Relator(a)                   |
| G      | Preferência/Sust. Oral       |
| H      | Modalidade                   |
| I      | Cabimento                    |
| J      | Advogado(a)                  |
| K      | OAB                          |
| O      | Parte Representada           |
| P      | 1º Polo Ativo                |
| Q      | 1º Polo Passivo              |

## Regras de ordenação automática

1. **Pedido de vista** — primeiro da lista
2. **Preferência legal** (idoso/gestante/lactante) — mesmo que videoconferência
3. **Presencial** — indicado "(PRESENCIAL)" após a OAB
4. **Videoconferência** — sem indicação (presume-se vídeo)

## Regras de concordância

- PELO/PELA conforme gênero da parte
- Empresas (Ltda, S.A.) = feminino
- Entidades masculinas (Conselho, Instituto, Banco, Condomínio, Município, BNDES) = masculino
- Pessoas = conforme nome
- APELADO/APELADA conforme gênero do polo passivo
- Plural (e OUTRO/OUTRA) = PELOS/PELAS + APELANTES/APELADOS

## Requisitos

- Navegador moderno (Chrome, Edge, Firefox)
- Zero instalação necessária
- Funciona 100% offline

## Arquivos

```
📁 Previa-Sustentacao-Oral/
  📄 index.html        — Interface + lógica
  📄 docx.iife.js      — Biblioteca de geração Word
  📄 jszip.min.js       — Biblioteca de leitura Excel
  📄 LEIA-ME.md         — Este arquivo
```
