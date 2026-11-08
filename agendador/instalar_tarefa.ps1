# Registra a tarefa diaria que roda o relatorio oculto (sem janela de console) no Agendador de Tarefas do Windows.
#   .\agendador\instalar_tarefa.ps1 -Hora 07:00
param([string]$Hora = '07:00', [string]$Nome = 'Relatorio - Titulos a pagar')
$vbs = Join-Path $PSScriptRoot 'executar_oculto.vbs'
$acao = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbs`""
$gatilho = New-ScheduledTaskTrigger -Daily -At $Hora
$config = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
Register-ScheduledTask -TaskName $Nome -Action $acao -Trigger $gatilho -Settings $config -Force `
    -Description 'Relatorio diario de titulos a pagar vencidos e a vencer'
