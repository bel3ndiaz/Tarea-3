// Editor de pasos de Registrar y Modificar: agregar, quitar y reordenar.
// Solo mueve filas en la pantalla; los pasos se guardan al enviar el formulario.
(function () {
  var lista = document.getElementById("lista-pasos");
  if (!lista) { return; }
  var plantilla = document.getElementById("plantilla-paso");

  // Vuelve a escribir los números 1, 2, 3... según el orden actual
  function renumerar() {
    var pasos = lista.querySelectorAll(".paso");
    for (var i = 0; i < pasos.length; i++) {
      pasos[i].querySelector(".paso-numero").textContent = i + 1;
    }
  }

  document.getElementById("agregar-paso").addEventListener("click", function () {
    lista.appendChild(plantilla.content.cloneNode(true));
    renumerar();
    lista.lastElementChild.querySelector("textarea").focus();
  });

  lista.addEventListener("click", function (evento) {
    var boton = evento.target.closest("button");
    if (!boton) { return; }
    var paso = boton.closest(".paso");
    if (boton.classList.contains("quitar")) {
      if (lista.children.length > 1) {
        paso.remove();
      } else {
        // Siempre queda al menos una fila: solo se vacía
        paso.querySelector("textarea").value = "";
        paso.querySelector("input").value = "";
      }
    } else if (boton.classList.contains("mover-arriba") && paso.previousElementSibling) {
      lista.insertBefore(paso, paso.previousElementSibling);
    } else if (boton.classList.contains("mover-abajo") && paso.nextElementSibling) {
      lista.insertBefore(paso.nextElementSibling, paso);
    }
    renumerar();
  });
})();
