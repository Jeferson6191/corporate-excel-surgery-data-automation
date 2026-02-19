<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    Automação e Padronização de Planilhas de Cirurgias
</head>
<body>

<h1>🏥 Automação e Padronização de Planilhas de Cirurgias</h1>
<h3>Excel + Pandas (Python)</h3>

<hr>

<h2>📌 Visão Geral</h2>
<p>
Este projeto consiste em uma automação desenvolvida em <strong>Python</strong> para o tratamento,
padronização e reorganização de planilhas Excel contendo dados de cirurgias hospitalares.
</p>

<p>
O objetivo principal é eliminar o trabalho manual diário de ajuste dessas planilhas,
garantindo consistência nos dados para posterior consumo em uma base analítica no
<strong>Power BI</strong>.
</p>

<p>
<strong>Projeto desenvolvido em contexto corporativo real</strong>, adaptado para fins de portfólio,
sem exposição de dados sensíveis.
</p>

<hr>

<h2>🧠 Problema</h2>
<p>
As planilhas continham registros de cirurgias que precisavam ser tratados manualmente
todos os dias, principalmente no campo <strong>"Setor cirurgia"</strong>, com base em regras específicas:
</p>

<ul>
    <li>Tipo da cirurgia</li>
    <li>Status do procedimento</li>
    <li>Procedimentos obstétricos (parto / cesariana)</li>
    <li>Setores cirúrgicos específicos (Centro Obstétrico, Centro Cirúrgico, Hemodinâmica)</li>
    <li>Regras diferentes dependendo da unidade hospitalar</li>
</ul>

<p>Esse processo manual era:</p>
<ul>
    <li>Repetitivo</li>
    <li>Suscetível a erro humano</li>
    <li>Pouco escalável</li>
    <li>Inviável para análises confiáveis</li>
</ul>

<hr>

<h2>🎯 Objetivo da Automação</h2>
<ul>
    <li>Padronizar automaticamente o campo <strong>Setor cirurgia</strong></li>
    <li>Aplicar regras de negócio diferentes por arquivo/unidade</li>
    <li>Filtrar procedimentos com base em critérios clínicos</li>
    <li>Remover colunas desnecessárias</li>
    <li>Gerar arquivos finais prontos para integração com Power BI</li>
</ul>

<hr>

<h2>🛠️ Tecnologias Utilizadas</h2>
<ul>
    <li>Python 3</li>
    <li>Pandas</li>
    <li>Excel (.xlsx)</li>
    <li>OS / Pathlib</li>
</ul>

<hr>

<h2>📂 Estrutura do Projeto</h2>
<pre>
excel-automation/
 ├─ main.py
 ├─ requirements.txt
 └─ README.md
</pre>

<hr>

<h2>⚙️ Regras de Negócio Implementadas</h2>

<h3>1️⃣ Cirurgias Obstétricas</h3>
<p>
Procedimentos como parto e cesariana são automaticamente classificados como
<strong>Centro Obstétrico</strong>, desde que:
</p>
<ul>
    <li>Tipo = Cirurgia</li>
    <li>Status = Realizada</li>
</ul>

<h3>2️⃣ Cirurgias Não Obstétricas</h3>
<p>
Cirurgias realizadas que não pertencem à lista de procedimentos obstétricos são
redirecionadas para <strong>Centro Cirúrgico</strong>.
</p>

<h3>3️⃣ Hemodinâmica</h3>
<p>
Cirurgias cujo setor inicia com <code>Hemodin</code> são padronizadas como
<strong>Hemodinâmica</strong>, quando aplicável à unidade hospitalar.
</p>

<h3>4️⃣ Outros Setores</h3>
<p>
Registros que não se enquadram nos setores esperados são ajustados automaticamente
para evitar inconsistências na base analítica.
</p>

<hr>

<h2>🔄 Fluxo de Processamento</h2>
<ol>
    <li>Leitura dos arquivos Excel</li>
    <li>Aplicação de filtros com Pandas</li>
    <li>Padronização do campo Setor cirurgia</li>
    <li>Remoção de colunas irrelevantes</li>
    <li>Geração de novos arquivos tratados (<code>T_nome_do_arquivo.xlsx</code>)</li>
</ol>

<hr>

<h2>▶️ Como Executar</h2>

<h3>Pré-requisitos</h3>
<ul>
    <li>Python 3 instalado</li>
</ul>

<h3>Instalação das dependências</h3>
<pre>
pip install pandas openpyxl
</pre>

<h3>Execução</h3>
<pre>
python main.py
</pre>

<p>
Os arquivos tratados serão gerados automaticamente na mesma pasta configurada.
</p>

<hr>

<h2>🔒 Observações Importantes</h2>
<ul>
    <li>Os nomes das unidades hospitalares foram mantidos apenas como referência técnica</li>
    <li>Nenhuma informação sensível ou confidencial é compartilhada</li>
    <li>Projeto adaptado para portfólio mantendo a lógica original de negócio</li>
</ul>

<hr>

<h2>👨‍💻 Autor</h2>
<p>
Projeto desenvolvido por <strong>Jefeferson Rodrigues</strong>, com foco em automação de processos,
tratamento e padronização de dados utilizando Python e Pandas para apoio a
análises e soluções de Business Intelligence.
</p>

</p>

</body>
</html>
