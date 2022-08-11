var prev_radio_checked = 'csv';
function op_change(input){
    op = ['a', 'b', 'c', 'd', 'e'].indexOf(input.value) >= 0 ? prev_radio_checked : input.value;
    if (op != 'f') { prev_radio_checked = op; }
   
    let all_a = document.querySelectorAll('section.no-border > article');
    for (let a of all_a){ a.style.display = 'none'; }

    let visibles_a = document.getElementsByClassName(op);
    for (let a of visibles_a){ a.style.display = 'inline-block'; }
}


function gerar_relatorio(){
    if (!document.getElementById('id_caminho').innerText){
        alert('Selecione um caminho para a saída');
        return false;
    }
    let args = [];
    let inputs = document.getElementsByTagName('input');
    for(let radio of inputs){
        if(radio.checked){
            args.push(radio.value);
        }
    }
    
    args = args[0] == 'f'  ? [args[0], 'csv', 'nao_separa', 'nao_tratado'] : args;
    args = args[1] == 'py' ? [args[0], args[1], 'nao_separa', 'tratado']   : args;
    args[0] == 'f' ? args.push('dp') : args.push('ae');

    eel.run(args)(result => result ? alert('Concluído') : alert('Erro ao gerar planilhas'));
}


function selecionar_pasta(){
    eel.busca()(result => document.getElementById('id_caminho').innerText = result);
}