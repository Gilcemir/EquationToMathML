# EquationToMathML
Macro para o Word, transformar as Equações do Word em MathMl, sem perder a tag "mml:"

Adaptei o código do projeto: https://github.com/CommWebTeam/vba/blob/748ac5ebc45c3bf801e2df7c6762e2e203443847/README.md

## Instalação

 - Adicione a aba de developer do Word
 - Vá na Aba "Visual Basic" 
 - Adicione o arquivo FM20.DLL em Ferramenta -> referências -> procurar em System32. Cheque se "Microsoft Forms 2.0 Object Library" está ativo.
 - Feche e abra de novo o word

Para rodar a macro basta clicar e fim!

OBS.: Para a macro funcionar, lembrar que a opção de copiar as fórmulas para o formato MathML tem que estar ativa dentro do Word.

## Edit:
 Adicionei mais duas Macros. 
 A primeira substitui o início da fórmula com a tag de equation, incrementando em cada fórmula a id da mesma. e1, e2, e3 etc.
 A segunda substitui e insere a tag equation no final.


Thanks to @CommWebTeam !
