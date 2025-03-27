document.addEventListener("DOMContentLoaded", function () {
    const chatList = document.getElementById("chat-list");
    const fetchButton = document.getElementById("fetch-chats");

    fetchButton.addEventListener("click", async function () {
        chatList.innerHTML = "Loading chats...";

        try {
            const response = await fetch("http://localhost:3000/getChats");
            const data = await response.json();

            if (data.error) {
                chatList.innerHTML = "Error fetching chats.";
                return;
            }

            chatList.innerHTML = ""; // Clear previous results

            // Display chat messages
            data.value.forEach(chat => {
                const li = document.createElement("li");
                li.textContent = chat.subject || "No subject";
                chatList.appendChild(li);
            });
        } catch (error) {
            console.error("Error fetching chats:", error);
            chatList.innerHTML = "Failed to load chats.";
        }
    });
});
