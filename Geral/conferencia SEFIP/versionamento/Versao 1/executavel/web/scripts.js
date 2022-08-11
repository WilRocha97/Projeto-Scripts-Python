function obtem_dados() {
    eel.busca_dados()(r => {
        document.getElementById('id_caminho').value = r
    });
}

function inicia_rotina() {
    caminho = document.getElementById('id_caminho').value;
    var msg = document.getElementById('msg_saida')
    msg.innerText = 'Analisando Relatórios'
    eel.iniciar(caminho)(r => {
        if(r) {
            msg.innerText = 'Programa Finalizado com Sucesso.'
        } else {
            msg.innerText = 'Programa Finalizado com Erro.'
        }
    })
}

function ajusta_tamanho(){
    window.resizeTo(600, 190)
}


window.onload = function() {
    document.addEventListener("contextmenu", function(e){
        e.preventDefault();
    }, false);
    document.addEventListener("keydown", function(e) {
    //document.onkeydown = function(e) {
        // "F12" key
        if (event.keyCode == 123) {
            disabledEvent(e);
        }
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