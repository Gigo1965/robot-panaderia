/* SaaS de demostracion: confirmar antes de escribir y avisar mientras se guarda. */
document.addEventListener('submit', function (evento) {
  var formulario = evento.target;
  var pregunta = formulario.getAttribute('data-confirmar');
  if (pregunta && !window.confirm(pregunta)) {
    evento.preventDefault();
    return;
  }
  if (formulario.hasAttribute('data-escribe')) {
    var boton = formulario.querySelector('button[type=submit]');
    if (boton) { boton.textContent = 'Guardando en el Excel…'; }
    document.querySelectorAll('form[data-escribe] button').forEach(function (b) { b.disabled = true; });
    var aviso = document.getElementById('guardando');
    if (aviso) { aviso.hidden = false; aviso.scrollIntoView({ block: 'nearest' }); }
  }
});
/* Si la persona vuelve con el boton Atras, los botones se reactivan. */
window.addEventListener('pageshow', function (evento) {
  if (evento.persisted) { window.location.reload(); }
});
