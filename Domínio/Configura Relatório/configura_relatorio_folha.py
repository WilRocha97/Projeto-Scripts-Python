# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, get_comp
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login


def configura_provisao_ferias():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('p')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('f')
    
    while not _find_img('provisao_ferias.png', conf=0.9):
        time.sleep(1)
    
    botoes = [('encargo_total.png', 'encargo_total_marcado.png'),
                    ('destacar_linhas.png', 'destacar_linhas_marcado.png'),
                    ('nao_considera_demitidos.png', 'nao_considera_demitidos_marcado.png'),]
    
    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('provisao_ferias.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Férias')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Férias')
  
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Provisão de Férias')
            print('❌ Erro ao procurar o botão')

    return True


def configura_provisao_decimo():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('p')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('d')
    
    while not _find_img('provisao_decimo.png', conf=0.9):
        time.sleep(1)

    if _find_img('analitico.png', conf=0.99):
        _click_img('analitico.png', conf=0.99)
    time.sleep(1)
    
    if _find_img('encargo_total.png', conf=0.99):
        _click_img('encargo_total.png', conf=0.99)
        _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Décimo Terceiro Salário')
        
    elif _find_img('encargo_total_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Décimo Terceiro Salário')
        
    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Provisão de Décimo Terceiro Salário')
        print('❌ Erro ao procurar o botão')
    time.sleep(1)
    
    if _find_img('completo.png', conf=0.99):
        _click_img('completo.png', conf=0.99)
    time.sleep(1)
    
    botoes = [('descontar_13_pago_rescisao.png', 'descontar_13_pago_rescisao_marcado.png'),
              ('descontar_adiantamento_13.png', 'descontar_adiantamento_13_marcado.png'),
              ('destacar_linhas.png', 'destacar_linhas_marcado.png'),
              ('nao_considerar_demitidos_mes.png', 'nao_considerar_demitidos_mes_marcado.png')]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('provisao_decimo.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Décimo Terceiro Salário')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Provisão de Décimo Terceiro Salário')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Provisão de Décimo Terceiro Salário')
            print('❌ Erro ao procurar o botão')
    

def configura_relatorios_admissionais():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('c')
    time.sleep(0.5)
    p.press('m')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('relatorios_admissionais.png', conf=0.9):
        time.sleep(1)

    _click_img('relatorios_adm.png', conf=0.99)
    time.sleep(1)

    botoes = [('acordo_compen_h_trabalhadas.png', 'acordo_compen_h_trabalhadas_marcado.png'),
              ('acordo_prorrog_h_trabalhadas.png', 'acordo_prorrog_h_trabalhadas_marcado.png'),
              ('autorizacao_desconto.png', 'autorizacao_desconto_marcado.png'),
              ('autorizacao_contr_sindical.png', 'autorizacao_contr_sindical_marcado.png'),
              ('contrato_experiencia.png', 'contrato_experiencia_marcado.png'),
              ('contrato_prazo_indeterminado.png', 'contrato_prazo_indeterminado_marcado.png'),
              ('declaracao_desl_vale_transporte.png', 'declaracao_desl_vale_transporte_marcado.png'),
              ('declaracao_enc_famil_fins_ir.png', 'declaracao_enc_famil_fins_ir_marcado.png'),
              ('declaracao_ren_vale_transporte.png', 'declaracao_ren_vale_transporte_marcado.png'),
              ('ficha_sal_familia.png', 'ficha_sal_familia_marcado.png'),
              ('termo_resp_sal_familia.png', 'termo_resp_sal_familia_marcado.png'),
              ('receb_devol_ctps.png', 'receb_devol_ctps_marcado.png'),
              ('ficha_registro_empregado.png', 'ficha_registro_empregado_marcado.png'),
              ('declaracao_opc_fgts.png', 'declaracao_opc_fgts_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('relatorios_admissionais.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Relatórios Admissionais')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Relatórios Admissionais')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Relatórios Admissionais')
            print('❌ Erro ao procurar o botão')


def configura_ficha_de_empregado():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('c')
    time.sleep(0.5)
    p.press('f')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('ficha_de_empregado.png', conf=0.9):
        time.sleep(1)

    _click_img('opcoes.png', conf=0.99)
    time.sleep(1)

    botoes = [('emitir_folha_anexa_hist_n_list_ficha_principal.png', 'emitir_folha_anexa_hist_n_list_ficha_principal_marcado.png'),
              ('demonstrar_info_contrato_prazo_determinado.png', 'demonstrar_info_contrato_prazo_determinado_marcado.png'),
              ('demonstrar_grade_salarial_prof.png', 'demonstrar_grade_salarial_prof_marcado.png'),
              ('demonstrar_altera_grade_salarial.png', 'demonstrar_altera_grade_salarial_marcado.png'),
              ('discr_h_semanal_empregado.png', 'discr_h_semanal_empregado_marcado.png'),
              ('discr_trocas_h_empregado.png', 'discr_trocas_h_empregado_marcado.png'),
              ('emitir_info_adv_susp.png', 'emitir_info_adv_susp_marcado.png'),
              ('demonstrar_info_ref_treina_capaci.png', 'demonstrar_info_ref_treina_capaci_marcado.png'),
              ('emitir_info_aviso_previo_indenizado.png', 'emitir_info_aviso_previo_indenizado_marcado.png'),
              ('emitir_info_ant_sal_hist_alt_salariais.png', 'emitir_info_ant_sal_hist_alt_salariais_marcado.png'),
              ('emitir_motivo_alt_salarial.png', 'emitir_motivo_alt_salarial_marcado.png'),
              ('demonstrar_data_afastamento.png', 'demonstrar_data_afastamento_marcado.png'),
              ('nao_emitir_dados_art8_port41_2007.png', 'nao_emitir_dados_art8_port41_2007_marcado.png'),
              ('espaco_ass_empregador.png', 'espaco_ass_empregador_marcado.png'),
              ('emitir_h_trab_colab_turno_vari_esp_esc.png', 'emitir_h_trab_colab_turno_vari_esp_esc_marcado.png'),
              ('gera_rela_quebra_empregado.png', 'gera_rela_quebra_empregado_marcado.png'),
              ('demonstrar_info_rne.png', 'demonstrar_info_rne_marcado.png'),
              ('demonstrar_info_serv_vinc_empregado.png', 'demonstrar_info_serv_vinc_empregado_marcado.png'),
              ('demonstrar_info_campo_livre.png', 'demonstrar_info_campo_livre_marcado.png'),
              ('emitir_ficha_mestra.png', 'emitir_ficha_mestra_marcado.png'),
              ('utilizar_nome_social_colaborador.png', 'utilizar_nome_social_colaborador_marcado.png'),
              ('emitir_data_hora.png', 'emitir_data_hora_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('ficha_de_empregado.png', conf=0.99)
            _escreve_relatorio_csv('Ok', end=';', nome='Ficha de empregado')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Ficha de empregado')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Ficha de empregado')
            print('❌ Erro ao procurar o botão')
    
    
def configura_pagamento_folha():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('m')
    time.sleep(0.5)
    p.press('f')
    
    while not _find_img('pagamento_de_folha.png', conf=0.9):
        time.sleep(1)

    if _find_img('participacao_lucros.png', conf=0.99):
        _click_img('participacao_lucros.png', conf=0.99)
        _click_img('pagamento_de_folha.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Pagamento de Folha')
        
    elif _find_img('participacao_lucros_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Pagamento de Folha')
        
    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Pagamento de Folha')
        print('❌ Erro ao procurar o botão')
    
    
def configura_requerimento_seguro_desemprego():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('s')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('requerimento_seguro_desemprego.png', conf=0.9):
        time.sleep(1)

    botoes = [('considerar_med_add_de_ferias.png', 'considerar_med_add_de_ferias_marcado.png'),
              ('considerar_add_de_ferias.png', 'considerar_add_de_ferias_marcado.png'),
              ('considerar_desc_faltas.png', 'considerar_desc_faltas_marcado.png'),
              ('considerar_sal_contrat_meses_susp_reduc.png', 'considerar_sal_contrat_meses_susp_reduc_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('requerimento_seguro_desemprego.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Requerimento de Seguro-Desemprego')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Requerimento de Seguro-Desemprego')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Requerimento de Seguro-Desemprego')
            print('❌ Erro ao procurar o botão')
    
    
def configura_integracao_contabil():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('integracao_contabil.png', conf=0.9):
        time.sleep(1)

    botoes = [('emitir_valores_prov_ferias.png', 'emitir_valores_prov_ferias_marcado.png'),
              ('emitir_valores_prov_13.png', 'emitir_valores_prov_13_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('integracao_contabil.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Integração Contabil')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Integração Contabil')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Integração Contabil')
            print('❌ Erro ao procurar o botão')


def configura_condicao_dif_trabalho():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('condicao_dif_trabalho.png', conf=0.9):
        time.sleep(1)
    
    if _find_img('emitir_info_requisitos_epi.png', conf=0.99):
        _click_img('emitir_info_requisitos_epi.png', conf=0.99)
        _click_img('condicao_dif_trabalho.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Condição Diferênciada de Trabalho')
        
    elif _find_img('emitir_info_requisitos_epi_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Condição Diferênciada de Trabalho')
        
    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Condição Diferênciada de Trabalho')
        print('❌ Erro ao procurar o botão')
    
    
def configura_quadro_h_de_trabalho():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('q')
    
    while not _find_img('quadro_h_trabalho.png', conf=0.9):
        time.sleep(1)
        
    if _find_img('detalhar_h_trabalho.png', conf=0.99):
        _click_img('detalhar_h_trabalho.png', conf=0.99)
        _click_img('quadro_h_trabalho.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Quadro de Horário de Trabalho')
        
    elif _find_img('detalhar_h_trabalho_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Quadro de Horário de Trabalho')

    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Quadro de Horário de Trabalho')
        print('❌ Erro ao procurar o botão')
    

def configura_etiquetas_ctps():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('c')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('etiquetas_ctps.png', conf=0.9):
        time.sleep(1)

    botoes = [('contrato_trabalho.png', 'contrato_trabalho_marcado.png'),
              ('contrato_xp.png', 'contrato_xp_marcado.png'),
              ('alteracao_salario.png', 'alteracao_salario_marcado.png'),
              ('anotacoes_ferias.png', 'anotacoes_ferias_marcado.png'),
              ('alteracoes_cargo.png', 'alteracoes_cargo_marcado.png'),
              ('livre.png', 'livre_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('etiquetas_ctps.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Etiquetas para CTPS')

        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Etiquetas para CTPS')

        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Etiquetas para CTPS')
            print('❌ Erro ao procurar o botão')
    

def configura_contrato_prazo_determinado():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('o')
    time.sleep(0.5)
    p.press('c')
    time.sleep(0.5)
    p.press('c')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('contrato_prazo_determinado.png', conf=0.9):
        time.sleep(1)
        
    if _find_img('contrato_de_xp.png', conf=0.99):
        _click_img('contrato_de_xp.png', conf=0.99)
        _click_img('contrato_prazo_determinado.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Contrato por Prazo Determinado')

    elif _find_img('contrato_de_xp_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Contrato por Prazo Determinado')

    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Contrato por Prazo Determinado')
        print('❌ Erro ao procurar o botão')
    
 
def configura_gfip():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('i')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('m')
    time.sleep(0.5)
    p.press('g')
    
    while not _find_img('gfip.png', conf=0.9):
        time.sleep(1)
    
    if _find_img('incluir_info_end_trabalhador.png', conf=0.99):
        _click_img('incluir_info_end_trabalhador.png', conf=0.99)
        _click_img('gfip.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Informativo Mensal GFIP')
        
    elif _find_img('incluir_info_end_trabalhador_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Informativo Mensal GFIP')
        
    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Informativo Mensal GFIP')
        print('❌ Erro ao procurar o botão')


def configura_comprovante_rendimentos():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('i')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('a')
    time.sleep(0.5)
    p.press('c')
    
    while not _find_img('comprovante_rendimentos.png', conf=0.9):
        time.sleep(1)

    botoes = [('consolidar_rend_matriz_filial.png', 'consolidar_rend_matriz_filial_marcado.png'),
              ('dispensada_ass.png', 'dispensada_ass_marcado.png'),
              ('gerar_relat_quebra_empregado.png', 'gerar_relat_quebra_empregado_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('comprovante_rendimentos.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Comprovante de Rendimentos')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Comprovante de Rendimentos')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Comprovante de Rendimentos')
            print('❌ Erro ao procurar o botão')


def configura_dirf_2009():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('i')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('a')
    time.sleep(0.5)
    p.press('d')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)

    botoes = [('declaracao_extincao.png', 'declaracao_extincao_marcado.png'),
              ('declarar_empregados.png', 'declarar_empregados_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('dirf.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='DIRF até 2009')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='DIRF até 2009')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='DIRF até 2009')
            print('❌ Erro ao procurar o botão')


def configura_gps():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('g')
    time.sleep(0.5)
    p.press('i')
    time.sleep(0.5)
    p.press('g')
    
    while not _find_img('gps.png', conf=0.9):
        time.sleep(1)

    botoes = [('preencher_total.png', 'preencher_total_marcado.png'),
              ('emitir_cod_empresa.png', 'emitir_cod_empresa_marcado.png'),]

    for botao in botoes:
        if _find_img(botao[0], conf=0.99):
            _click_img(botao[0], conf=0.99)
            _click_img('gps.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='GPS')
            
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='GPS')
            
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='GPS')
            print('❌ Erro ao procurar o botão')


def configura_grcsu_patronal():
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    p.press('g')
    time.sleep(0.5)
    p.press('t')
    
    while not _find_img('grcsu_patronal.png', conf=0.9):
        time.sleep(1)
        
    if _find_img('demonstrar_cod_ent_sind_completo.png', conf=0.99):
        _click_img('demonstrar_cod_ent_sind_completo.png', conf=0.99)
        _click_img('grcsu_patronal.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='GRCSU Patronal')
        
    elif _find_img('demonstrar_cod_ent_sind_completo_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='GRCSU Patronal')
        
    else:
        _escreve_relatorio_csv('Erro', end=';', nome='GRCSU Patronal')
        print('❌ Erro ao procurar o botão')


def configura_avisos_vencimento():
    p.hotkey('alt', 'c')
    time.sleep(0.5)
    p.press('v')
    
    while not _find_img('avisos_vencimento.png', conf=0.9):
        time.sleep(1)

    botoes = [('competencia_sem_fechamento_esocial.png', 'competencia_sem_fechamento_esocial_marcado.png', 'competencia_sem_fechamento_esocial_nome.png'),
              ('envio_rescisao_esocial.png', 'envio_rescisao_esocial_marcado.png', 'envio_rescisao_esocial_nome.png'),
              ('atestado_frequencia.png', 'atestado_frequencia_marcado.png', 'atestado_frequencia_nome.png'),
              ('carteira_vacinacao.png', 'carteira_vacinacao_marcado.png', 'carteira_vacinacao_nome.png'),
              ('carteira_motorista.png', 'carteira_motorista_marcado.png', 'carteira_motorista_nome.png'),
              ('vencimento_epi.png', 'vencimento_epi_marcado.png', 'vencimento_epi_nome.png'),
              ('vencimento_2_ferias.png', 'vencimento_2_ferias_marcado.png', 'vencimento_2_ferias_nome.png'),
              ('vencimento_epi.png', 'vencimento_epi_marcado.png', 'vencimento_epi_nome.png'),]

    for botao in botoes:
        while not _find_img(botao[2], conf=0.9):
            time.sleep(1)
            _click_img('down.png', conf=0.9)
            _click_img('avisos_vencimento.png', conf=0.9)
            
        if _find_img(botao[0], conf=0.99):
            
            p.moveTo(p.locateCenterOnScreen(os.path.join('imgs', botao[0]), confidence=0.9))
            local_mouse = p.position()
            p.click(int(local_mouse[0] + 110), local_mouse[1])
            time.sleep(0.2)
            p.press('up')
            time.sleep(0.2)
            p.press('enter')
            
            _click_img('avisos_vencimento.png', conf=0.9)
            _escreve_relatorio_csv('Ok', end=';', nome='Avisos de Vencimento')
    
        elif _find_img(botao[1], conf=0.99):
            _escreve_relatorio_csv('Ok', end=';', nome='Avisos de Vencimento')
    
        else:
            _escreve_relatorio_csv('Erro', end=';', nome='Avisos de Vencimento')
            print('❌ Erro ao procurar o botão')

    p.hotkey('alt', 'g')


def configura_parametros():
    p.hotkey('alt', 'c')
    time.sleep(0.5)
    p.press('p')
    
    while not _find_img('parametros.png', conf=0.9):
        time.sleep(1)
    
    _click_img('sst.png', conf=0.9)
    
    while not _find_img('opcoes_sst.png', conf=0.9):
        time.sleep(1)

    if _find_img('nao_envia_eventos_sst_apartir_de.png', conf=0.99):
        _click_img('nao_envia_eventos_sst_apartir_de.png', conf=0.99)
        _click_img('opcoes_sst.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')

    elif _find_img('nao_envia_eventos_sst_apartir_de_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')

    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Parâmetros')
        print('❌ Erro ao procurar o botão')

    _click_img('mais_opcoes.png', conf=0.99)
    
    while not _find_img('eventos_sst.png', conf=0.9):
        time.sleep(1)
    
    p.press('tab')
    time.sleep(1)
    if _find_img('com_acid_trab.png', conf=0.99):
        _click_img('com_acid_trab.png', conf=0.99)
        _click_img('eventos_sst.png', conf=0.9)
        _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')

    elif _find_img('com_acid_trab_marcado.png', conf=0.99):
        _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')

    else:
        _escreve_relatorio_csv('Erro', end=';', nome='Parâmetros')
        print('❌ Erro ao procurar o botão')
        
    time.sleep(1)
    p.hotkey('alt', 'f')

    botoes_personaliza = [('considerar_dias_calc_ini_estagio.png', 'considerar_dias_calc_ini_estagio_marcado.png'),
                     ('permitir_info_anali_sinte_centro_custos.png', 'permitir_info_anali_sinte_centro_custos_marcado.png'),]

    botoes_dsr = [('descontar_faltas_dsr_conf_folga.png', 'descontar_faltas_dsr_conf_folga_marcado.png'),]

    botoes_salario_familia = [('nao_calc_empregado_inter_comp_n_calculo.png', 'nao_calc_empregado_inter_comp_n_calculo_marcado.png'),]

    botoes_rescisao = [('gerar_val_aviso_previo_na_sefip.png', 'gerar_val_aviso_previo_na_sefip_marcado.png'),
                     ('calcular_multa_art479e480.png', 'calcular_multa_art479e480_marcado.png'),
                     ('calcular_aviso_previo_acordo_metade.png', 'calcular_aviso_previo_acordo_metade_marcado.png'),]

    botoes_arredondamento = [('empregados.png', 'empregados_marcado.png'),
                       ('estagiarios.png', 'estagiarios_marcado.png'),
                       ('contribuintes.png', 'contribuintes_marcado.png'), ]

    botoes_adiantamento_definicoes = [('para_estagiarios.png', 'para_estagiarios_marcado.png'),
                             ('para_contribuintes.png', 'para_contribuintes_marcado.png'),
                             ('para_aprendiz.png', 'para_aprendiz_marcado.png'), ]

    botoes_adiantamento_proporcionalidade = [('considerar_dias_trabalhados.png', 'considerar_dias_trabalhados_marcado.png'),
                                      ('considerar_dias_trabalhados.png', 'considerar_dias_trabalhados_marcado.png'),]

    botoes_ferias = [('descontar_faltas_suspensas.png', 'descontar_faltas_suspensas_marcado.png'),
                     ('nao_calcula_sal_fami_nas_ferias.png', 'nao_calcula_sal_fami_nas_ferias_marcado.png'), ]

    botoes_contabil = [('gerar_integracao_contabil_colaborador.png', 'gerar_integracao_contabil_colaborador_marcado.png'),
                    ('folha_mensal.png', 'folha_mensal_marcado.png'),
                    ('ferias.png', 'ferias_marcado.png'),
                    ('rescisao.png', 'rescisao_marcado.png'),
                    ('provisao_de_ferias.png', 'provisao_de_ferias_marcado.png'),
                    ('provisao_13.png', 'provisao_13_marcado.png'),]
    
    botoes_opcoes = [('integrar_irrf_ferias_data_programada.png', 'integrar_irrf_ferias_data_programada_marcado.png'),
              ('gerar_valor_prov_ajust_sefip_pag_encargos.png', 'gerar_valor_prov_ajust_sefip_pag_encargos_marcado.png'),
              ('considerar_ult_dia_mes.png', 'considerar_ult_dia_mes_marcado.png'),]
    
    abas = [('opcoes_personaliza.png', 'personaliza.png', botoes_personaliza, ''),
            ('opcoes_dsr.png', 'dsr.png', botoes_dsr, ''),
            ('opcoes_salario_familia.png', 'salario_familia.png', botoes_salario_familia, ''),
            ('opcoes_rescisao.png', 'aba_rescisao.png', botoes_rescisao, ''),
            ('opcoes_arredondamento.png', 'arredondamento.png', botoes_arredondamento, ''),
            ('opcoes_adiantamento_definicoes.png', 'adiantamento.png', botoes_adiantamento_definicoes, 'adiantamento_definicoes.png'),
            ('opcoes_adiantamento_proporcionalidade.png', 'adiantamento.png', botoes_adiantamento_proporcionalidade, 'adiantamento_proporcionalidade.png'),
            ('opcoes_ferias.png', 'aba_ferias.png', botoes_ferias, ''),
            ('opcoes_integracao.png', 'contabilidade.png', botoes_contabil, ''),
            ('opcoes_parametros_cont.png', 'opcoes.png', botoes_opcoes, ''),]
    
    for aba in abas:
        while not _find_img(aba[0], conf=0.9):
            _click_img(aba[1], conf=0.99)
            if aba[3] != '':
                _click_img(aba[3], conf=0.99, timeout=1)
            time.sleep(1)
        
        for botao in aba[2]:
            if aba[3] != '':
                _click_img(aba[3], conf=0.99, timeout=1)
                time.sleep(1)
            else:
                _click_img(aba[1], conf=0.99, timeout=1)
                
            time.sleep(1)
            count = 0
            while not _find_img(botao[0], conf=0.99):
                p.press('tab')
                count += 1
                if count > 5:
                    break
                    
            if _find_img(botao[0], conf=0.99):
                _click_img(botao[0], conf=0.99)
                _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')
                _click_img('parametros.png', conf=0.9)
                continue
                
            count = 0
            while not _find_img(botao[1], conf=0.99):
                p.press('tab')
                count += 1
                if count > 5:
                    break
                    
            if _find_img(botao[1], conf=0.99):
                _escreve_relatorio_csv('Ok', end=';', nome='Parâmetros')
        
            else:
                _escreve_relatorio_csv('Erro', end=';', nome='Parâmetros')
                print('❌ Erro ao procurar o botão')
            time.sleep(1)
            
    p.hotkey('alt', 'g')
    time.sleep(1)
    if _find_img('atencao.png', conf=0.9):
        p.hotkey('alt', 's')
    time.sleep(3)
    
            
@_time_execution
def run():
    empresas = _open_lista_dados()
    andamentos = 'Configura Relatórios'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]

    relatorios = [(configura_provisao_ferias, 'Provisão de Férias', 'CÓDIGO;CNPJ;NOME;'),
        (configura_provisao_decimo, 'Provisão de Décimo Terceiro Salário', 'CÓDIGO;CNPJ;NOME;'),
        (configura_relatorios_admissionais, 'Provisão de Férias', 'CÓDIGO;CNPJ;NOME;'),
        (configura_ficha_de_empregado, 'Ficha de empregado', 'CÓDIGO;CNPJ;NOME;'),
        (configura_pagamento_folha, 'Pagamento de Folha', 'CÓDIGO;CNPJ;NOME;'),
        (configura_requerimento_seguro_desemprego, 'Requerimento de Seguro-Desemprego', 'CÓDIGO;CNPJ;NOME;'),
        (configura_integracao_contabil, 'Integração Contabil', 'CÓDIGO;CNPJ;NOME;'),
        (configura_condicao_dif_trabalho, 'Condição Diferênciada de Trabalho', 'CÓDIGO;CNPJ;NOME;'),
        (configura_quadro_h_de_trabalho, 'Quadro de Horário de Trabalho', 'CÓDIGO;CNPJ;NOME;'),
        (configura_etiquetas_ctps, 'Etiquetas para CTPS', 'CÓDIGO;CNPJ;NOME;'),
        (configura_contrato_prazo_determinado, 'Contrato por Prazo Determinado', 'CÓDIGO;CNPJ;NOME;'),
        (configura_gfip, 'Informativo Mensal GFIP', 'CÓDIGO;CNPJ;NOME;'),
        (configura_comprovante_rendimentos, 'Comprovante de Rendimentos', 'CÓDIGO;CNPJ;NOME;'),
        (configura_dirf_2009, 'DIRF até 2009', 'CÓDIGO;CNPJ;NOME;'),
        (configura_gps, 'GPS', 'CÓDIGO;CNPJ;NOME;'),
        (configura_grcsu_patronal, 'GRCSU Patronal', 'CÓDIGO;CNPJ;NOME;'),
        (configura_avisos_vencimento, 'Avisos de Vencimento', 'CÓDIGO;CNPJ;NOME;'),
        (configura_parametros, 'Parâmetros', 'CÓDIGO;CNPJ;NOME;'), ]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cod, cnpj, nome = empresa
        _indice(count, total_empresas, empresa)
        
        if not _login(empresa, andamentos):
            continue
        
        for relatorio in relatorios:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome}', end=';', nome=relatorio[1])
            
            relatorio[0]()
            print(f'✔ {relatorio[1]}')

            _escreve_relatorio_csv('\n', end='', nome=relatorio[1])
            p.press('esc', presses=3, interval=0.5)
        
    for relatorio in relatorios:
        _escreve_header_csv(relatorio[2], nome=relatorio[1])
            
        
if __name__ == '__main__':
    run()
