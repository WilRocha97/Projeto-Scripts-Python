function render_options(op){
    data.models.forEach(function (model){
        if (model.id == op){ this_model = model; }
    });

    var op_tmpl = $.templates('#optionTemplate');
    var op_html = op_tmpl.render(this_model);
    $("#options_div").html(op_html).text();
    render_panes($('div#options_div button.active').text());
    set_click_option();
}

function render_panes(op){
    this_pane = false;
    $(".pane").remove()
    panes.forEach(function (pane){
        if (pane.name == op){ this_pane = pane; }
    });

    if (!this_pane) { return false }
    var pane_tmpl = $.templates("#paneTemplate");
    var pane_html = pane_tmpl.render(this_pane);
    $("#options_div").append(pane_html);
}

/*Definir option onclicks*/
function set_click_option(){
    for (elem of $("button.list-group-item")) {
        if (['op_csv', 'op_py'].indexOf(elem.id) >= 0 ){
            $("#" + elem.id).click(function (ev){
                render_panes(ev.target.value);
            });
        }
        if ('op_extra_arq' === elem.id){
            $("#" + elem.id).click(function (){
                eel.get_file()(result => $('#op_extra_arq').text(result));
            });
        }
    }
}

$(document).ready(function(){
    var tmpl = $.templates("#modelTemplate");
    var html = tmpl.render(data.models);
    $("#models_div").append(html);
    render_options("a");

    /*Definir btn onclicks*/
    $("#btn_caminho").click(function(){
        eel.busca()(result => $('#span_caminho').text(result));
    });

    $("#btn_gerar").click(function(){
        if( $("#span_caminho").text() == "Selecione o caminho para salvar arquivo"){
            alert('O caminho para salvar o arquivo é obrigatorio');
            return false
        }
        if( $("#op_extra_arq") == null || $("#op_extra_arq").text() == "Selecione o arquivo complementar" ){
            alert('O arquivo complementar é obrigatorio');
            return false
        }

        args = []
        for (i of $('.active')){ args.push(i.value); }
        args.push('ae');

        /*Tratamentos especiais*/
        args = args[1] == 'py' ? [args[0], args[1], 'nao_separa', 'tratado', 'ae']  : args;
        args = args[1] == 'g'  ? [args[0], 'csv', 'nao_separa', 'nao_tratado', 'dp'] : args;
        args = args[1] == 'f'  ? [args[0], 'csv', 'nao_separa', 'nao_tratado', 'pdf_dp'] : args;
        
        eel.run(args)(result => result ? alert('Concluído') : alert('Erro ao gerar planilhas'));
    });
});