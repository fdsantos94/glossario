# Compatibilidade Visual e DPI Scaling

## ⚠️ O Problema
O **DPI Scaling** (Escala de DPI) é um recurso do Windows que ajusta o tamanho de interfaces gráficas baseando-se na densidade de pixels (Dots Per Inch).

Identificamos que, em máquinas com configurações de DPI específicas, o **UserForm do Glossário sofre distorções ou recortes**, tornando o uso impossível ou desagradável.

## 🛠️ A Solução Implementada
Tentativas de correção universal via código falharam, pois o ajuste para uma máquina quebrava a visualização em outra. Adotamos uma **abordagem híbrida baseada em exceções**.

### Lógica de Funcionamento
1. **Banco de Dados:** Criamos a tabela `banco.schema.GlossarioUsuariosExcecao` para listar os logins de rede que sofrem com esse problema.
2. **Validação no VBA:** Ao iniciar, o sistema verifica se o usuário logado consta nesta lista.
   - **Se SIM:** Aplica um fator de correção dimensional no UserForm.
   - **Se NÃO:** Mantém a renderização padrão do Windows.

### ⚙️ Implementação Técnica
Para que o mecanismo funcione, os seguintes requisitos devem ser mantidos:

* **View de Dados:** A view `select * from vw_GlossarioUsuariosExcecao` deve ser consumida dentro do módulo `xpto`.
* **Dump de Dados:** A lista de usuários é despejada na **Coluna AD** da aba (Sheet) `Glossário`.
* **Atualização:** Sempre que o relatório for atualizado (via delivery automático ou manual), essa lista deve ser renovada para garantir que novos usuários com problemas sejam capturados.