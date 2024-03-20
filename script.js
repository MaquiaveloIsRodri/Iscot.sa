// Obtenemos los nodos del árbol
var nodes = document.querySelectorAll(".node");

// Creamos las flechas
for (var i = 0; i < nodes.length; i++) {
    var node = nodes[i];

    // Obtenemos el ID del nodo padre
    var parentId = node.getAttribute("id");

    // Obtenemos el nodo padre
    var parent = document.getElementById(parentId);

    // Si el nodo tiene hijos
    if (parent.children.length) {
        // Creamos una flecha entre el nodo padre y cada uno de sus hijos
        for (var j = 0; j < parent.children.length; j++) {
            var child = parent.children[j];

            // Creamos la flecha
            var arrow = document.createElement("div");
            arrow.className = "arrow";
            arrow.setAttribute("data-from", parentId);
            arrow.setAttribute("data-to", child.id);

            // Añadimos la flecha al nodo hijo
            child.appendChild(arrow);
        }
    }
}
