const cards = document.querySelectorAll(".workspace-card");

cards.forEach(function(card){

    card.addEventListener("click", function(){

        const workspaceType = this.dataset.type;

        document.getElementById("workspace_type").value = workspaceType;

        const modalElement = document.getElementById("workspaceModal");

        const modal = new bootstrap.Modal(modalElement);

        modal.show();

    });

});