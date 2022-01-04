# Randomizador de turmas

esse algoritmo permite dividir uma lista de alunos para uma certa quantidade de turmas.

**Lista de aluno base:** *(banco.xlsx)*

|      | Relação Turma 1 |
| ---- | ----------------- |
| 101  | Bob               |
| 103  | Sally             |
| 104  | Sue               |
| 105  | Jill              |
| 106  | Jon               |
| 171  | Ted               |
| 1271 | millon            |
| 1471 | cody              |
| 8171 | lola              |
| 945  | zoro              |
| 4171 | luffy             |
| 1171 | sanji             |


**Regras para a randomização:** (*rulesx.xlsx)*

| 4 Turmas  |           |           |           |           |
| --------- | --------- | --------- | --------- | --------- |
| Módulo 1 | Módulo 2 | Módulo 3 | Módulo 4 | Módulo 5 |
| 11        | 41        | 31        | 21        | 11        |
| 21        | 31        | 11        | 41        | 31        |
| 31        | 21        | 41        | 11        | 21        |
| 41        | 11        | 21        | 31        | 41        |
|           |           |           |           |           |
|           |           |           |           |           |
| 3 Turmas  |           |           |           |           |
| Módulo 1 | Módulo 2 | Módulo 3 | Módulo 4 | Módulo 5 |
| 11        | 31        | 21        | 11        | 21        |
| 21        | 11        | 31        | 21        | 11        |
| 31        | 21        | 11        | 31        | 31        |
|           |           |           |           |           |
|           |           |           |           |           |
| 2 Turmas  |           |           |           |           |
| Módulo 1 | Módulo 2 | Módulo 3 | Módulo 4 | Módulo 5 |
| 11        | 21        | 21        | 11        | 21        |
| 21        | 11        | 11        | 21        | 11        |
|           |           |           |           |           |


**Lista após as randomizações:**

| Matricula | Nome   |  | Modulo 1 |    | Modulo 2 |    | Modulo 3 |    | Modulo 4 |    | Modulo 5 |    |
| --------- | ------ | - | -------- | -- | -------- | -- | -------- | -- | -------- | -- | -------- | -- |
| 101       | Bob    |  | 945      | 11 | 101      | 31 | 101      | 21 | 101      | 11 | 101      | 21 |
| 103       | Sally  |  | 105      | 11 | 103      | 31 | 103      | 21 | 103      | 11 | 103      | 21 |
| 104       | Sue    |  | 1171     | 11 | 104      | 31 | 104      | 21 | 104      | 11 | 104      | 21 |
| 105       | Jill   |  | 1271     | 11 | 105      | 31 | 105      | 21 | 105      | 11 | 105      | 21 |
| 106       | Jon    |  | 8171     | 21 | 106      | 11 | 106      | 31 | 106      | 21 | 106      | 11 |
| 171       | Ted    |  | 171      | 21 | 171      | 11 | 171      | 31 | 171      | 21 | 171      | 11 |
| 1271      | millon |  | 103      | 21 | 1271     | 11 | 1271     | 31 | 1271     | 21 | 1271     | 11 |
| 1471      | cody   |  | 106      | 21 | 1471     | 11 | 1471     | 31 | 1471     | 21 | 1471     | 11 |
| 8171      | lola   |  | 4171     | 31 | 8171     | 21 | 8171     | 11 | 8171     | 31 | 8171     | 31 |
| 945       | zoro   |  | 1471     | 31 | 945      | 21 | 945      | 11 | 945      | 31 | 945      | 31 |
| 4171      | luffy  |  | 101      | 31 | 4171     | 21 | 4171     | 11 | 4171     | 31 | 4171     | 31 |
| 1171      | sanji  |  | 104      | 31 | 1171     | 21 | 1171     | 11 | 1171     | 31 | 1171     | 31 |


`![tutorial em gif](https://media.giphy.com/media/vFKqnCdLPNOKc/giphy.gif)`
