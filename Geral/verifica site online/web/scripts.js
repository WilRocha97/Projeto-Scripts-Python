
function inicia_rotina(){
    document.getElementById('msg_saida').innerText = 'Verificando Site';
    let url = document.getElementById('id_site').value;
    eel.verifica(url)(()=>{});
}

eel.expose(status);
function status(status){
    let saida = document.getElementById('msg_saida'),
        div = document.getElementById('id_sinal');
    if(status){
        saida.innerText = 'Site Online'; 
        div.style.backgroundColor = 'green';
    } else {
        saida.innerText = 'Site Offline';
        div.style.backgroundColor = 'red';
    } 

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