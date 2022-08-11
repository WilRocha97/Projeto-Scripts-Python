function obtem_dados() {
    eel.busca_dados()(r => {
        document.getElementById('id_caminho').value = r
    });
}

function inicia_rotina() {
    var relatorio = document.getElementById('id_caminho').value;
    var tipo = document.querySelector("input:checked").value;
    var msg = document.getElementById('msg_saida')
    if (!relatorio){
        msg.innerText = "Informe um arquivo pdf";
        return false;
    }

    msg.innerText = 'Analisando arquivo';
    eel.iniciar(relatorio, tipo)(r => {
        if(r) {
            msg.innerText = 'Programa finalizado com sucesso.'
            let t = '';
            if (r !== true){
                t = `Páginas não lidas: ${r}`;
            }
            document.getElementById('id_paginas').innerText = t;
        } else {
            msg.innerText = 'Programa finalizado com erro.'
        }
    })
}

function ajusta_tamanho(){
    window.resizeTo(600, 230)
}

window.onload = function() {
    document.addEventListener("contextmenu", function(e){
        e.preventDefault();
    }, false);
    document.addEventListener("keydown", function(e) {
        if (event.keyCode == 123) disabledEvent(e); // "F12" key
    }, false);
    function disabledEvent(e){
        if (e.stopPropagation){
            e.stopPropagation();
        } else if (window.event){
            window.event.cancelBubble = true;
        }
        e.preventDefault();
        return false;
    }
};