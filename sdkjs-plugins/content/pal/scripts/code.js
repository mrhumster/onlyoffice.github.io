(function (window, undefined) {
    window.Asc.plugin.init = function () {
        const elements = {
            searchInput: document.getElementById('search_input'),
            search: document.getElementById('search'),
            authForm: document.getElementById('auth'),
            btnRemoveKey: document.getElementById('btn_remove_key')
        }
        const handleSubmit = (e) => {
            const api_key = e.target.elements['api_key'].value;
            localStorage.setItem("x-api-key", api_key);
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'block';
        }

        const handleChange = (e) => {
            const query = e.target.value;
            const key = localStorage.getItem("x-api-key");
            fetch('https://base/api/search?' + new URLSearchParams({query: query}), {
                headers: {
                    Accept: "application/json",
                    "x-api-key": key
                }
            })
                .then((response) => console.log(response))
                .catch((error) => console.log(error))
        }

        const handleRemoveKey = () => {
            localStorage.removeItem("x-api-key")
            elements.authForm.style.display = 'block';
            elements.search.style.display = 'none';
        }

        const key = localStorage.getItem("x-api-key");
        if (key) {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'block';
        } else {
            elements.authForm.style.display = 'block';
            elements.search.style.display = 'none';
        }

        elements.authForm.onsubmit = handleSubmit;
        elements.btnRemoveKey.onclick = handleRemoveKey;
        elements.searchInput.oninput = handleChange;
    };
    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);