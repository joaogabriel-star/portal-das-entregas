# Medição da recuperação (retrieval) do catálogo

Scripts que medem, com número, se a busca de candidatas está achando a entrega
certa. Rodam a partir da raiz do projeto: `node ferramentas/medicao/<arquivo>`.

Eles **extraem as funções reais** de `portal-frontend.jsx` (não reimplementam),
então medem o que está no ar de verdade.

| Arquivo | O que mede | Resultado quando foi escrito |
|---|---|---|
| `recall-recuperacao.js` | recall com a atribuição derivada da própria entrega | recall@20 = 98,9% · @50 = 100% |
| `recall-parafrase.js` | recall quando o gestor usa **outras palavras** | 1 de 6 dentro de 80 |
| `ganho-expansao.js` | o que muda ao traduzir a frase para o vocabulário do catálogo antes de buscar | 1 de 6 → 6 de 6 |

## O que esses números querem dizer

O casamento é léxico. Quando o texto do gestor usa as palavras do catálogo, a
recuperação é praticamente perfeita. Quando ele parafraseia ("celebração de
ajustes" no lugar de "Contratos firmados"), a entrega certa não chega ao
modelo — e aí não há prompt que salve.

Por isso existe `expandirConsulta()`: a IA reescreve a frase no vocabulário do
catálogo antes da busca. Se a IA não estiver disponível, o fluxo segue no
léxico e a tela avisa "busca local, sem IA".

## Ressalva honesta sobre o 1 de 6

O gabarito de `recall-parafrase.js` cobra **uma** entrega específica por caso.
O catálogo costuma ter várias entregas defensáveis para a mesma frase, então o
número subestima a qualidade percebida: na tela, mesmo sem IA, a busca devolve
resultados aproveitáveis. Use o script para comparar versões (antes/depois),
não como nota absoluta.
