
function get_coord(){
    var x = document.getElementById('id_coord_X').value;
    var y = document.getElementById('id_coord_Y').value;
    var rgb = document.getElementById('id_rgb').value;

    return [x, y, rgb]
}

function atualiza_ponto(ponto){
    var result = get_coord();
    var x = result[0];
    var y = result[1];
    document.getElementById(ponto).innerText = ' '+x+' , '+y;
}

window.onload = function() {
    eel.get_position();
    document.addEventListener("contextmenu", function(e){
        e.preventDefault();
    }, false);
    document.addEventListener("keydown", function(e) {
        if ((event.ctrlKey && event.shiftKey && event.keyCode == 73) || // Ctrl+Shft+i
             (event.ctrlKey && event.shiftKey && event.keyCode == 67) || // Ctrl+Shft+c
             (event.ctrlKey && event.shiftKey && event.keyCode == 78)) { // Ctrl+Shft+n
            disabledEvent(e);
        } else if ((event.ctrlKey && event.keyCode == 67) || // Ctrl+c
                    (event.ctrlKey && event.keyCode == 78) || // Ctrl+n
                    (event.ctrlKey && event.keyCode == 80) || // Ctrl+p
                    (event.ctrlKey && event.keyCode == 84) || // Ctrl+t
                    (event.ctrlKey && event.keyCode == 85) || // Ctrl+u
                    (event.ctrlKey && event.keyCode == 82) || // Ctrl+r
                    (event.ctrlKey && event.keyCode == 83)) { // Ctrl+s
            disabledEvent(e);
        }
        if ((event.keyCode == 116) || // F5
            (event.keyCode == 122) || // F11
            (event.keyCode == 123)) { // F12
            disabledEvent(e);
        } else if (event.keyCode == 117) { // F6
            atualiza_ponto('id_ponto_1')
        } else if (event.keyCode == 118) { // F7
            atualiza_ponto('id_ponto_2')
        } else if (event.keyCode == 119) { // F8
            var x = document.getElementById('id_ponto_1').innerText;
            var y = document.getElementById('id_ponto_2').innerText;

            if (x === '' || y === '') {
                document.getElementById('id_alerta').innerText = 'Defina os pontos da imagem.'
            } else {
                eel.salvar_imagem(x, y)(r => {
                    if (r) {
                        document.getElementById('id_alerta').innerText = r + '  Salva.'
                    } else {
                        document.getElementById('id_alerta').innerText = 'Falha ao salvar imagem.'
                    }
                })
            }
        } else if (event.keyCode == 120) { // F9
            var result = get_coord();
            var x = result[0];
            var y = result[1];
            var rgb = result[2];
            document.getElementById('id_saida').value += '('+x+', '+y+') - '+rgb+'\n';
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

eel.expose(atualiza_coordenadas)
function atualiza_coordenadas(x, y, rgb){
    document.getElementById('id_coord_X').value = x;
    document.getElementById('id_coord_Y').value = y;
    document.getElementById('id_rgb').value = rgb;
}

function busca_pasta(){
    eel.busca_caminho()(r => document.getElementById('id_salvar_em').innerText = r);
}
