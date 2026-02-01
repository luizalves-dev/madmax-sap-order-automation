# madmax-sap-order-automation

Automação desenvolvida em **Excel VBA** para orquestrar a captura, validação e processamento de pedidos em uma **transação SAP customizada**.

O projeto ficou conhecido internamente como **Madmax** devido à sua capacidade de processar grandes volumes de dados enquanto lida com regras operacionais e exceções em tempo real.

## Visão Geral
A solução organiza dados de entrada no Excel, interage com transações SAP e grids ALV, valida informações, trata exceções operacionais automaticamente e executa correções antes da submissão dos pedidos.

O objetivo é reduzir digitação manual, erros humanos e retrabalho, aumentando a eficiência operacional e a confiabilidade dos dados.

## Principais Funcionalidades
- Navegação automatizada em transação SAP customizada  
- Leitura e manipulação de ALV Grid  
- Validação e organização de dados antes da submissão  
- Tratamento automático de exceções (itens inválidos, restrições operacionais, condições de pagamento)  
- Supressão e ajuste automático de itens inconsistentes  
- Geração de registros e trilha de auditoria no Excel   

## Estrutura do Repositório

madmax-sap-order-automation/
├─ src/
│  ├─ Madmax_Digitador.bas
│  ├─ Madmax_Limpeza.bas
├─ releases/
│  ├─ Madmax.xlsb
├─ README.md
└─ .gitignore


## Observações Importantes
- Este repositório contém **apenas módulos VBA exportados (.bas)**  
- Nenhuma credencial, captura de tela, nome de transação interna ou regra de negócio sensível foi incluída  
- O código foi abstraído para fins de portfólio e demonstração técnica  

## Tecnologias Utilizadas
- Excel VBA  
- SAP GUI Scripting  
- ALV Grid  

---

Projeto com foco em **automação de processos, eficiência operacional e confiabilidade de dados** em ambiente corporativo.




