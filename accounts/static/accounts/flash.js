document.addEventListener('DOMContentLoaded', () => {
    const messages = document.querySelectorAll('[data-flash-message]');

    const dismissMessage = (message) => {
        if (!message || message.classList.contains('is-leaving')) {
            return;
        }

        message.classList.add('is-leaving');
        window.setTimeout(() => {
            message.remove();
        }, 320);
    };

    messages.forEach((message) => {
        const closeButton = message.querySelector('[data-flash-close]');

        if (closeButton) {
            closeButton.addEventListener('click', () => dismissMessage(message));
        }

        window.setTimeout(() => dismissMessage(message), 3600);
    });
});
