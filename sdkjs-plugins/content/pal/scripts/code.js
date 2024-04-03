function debounce(func, wait, immediate) {
    var timeout;
    return function() {
        var context = this, args = arguments;
        var later = function() {
            timeout = null;
            if (!immediate) func.apply(context, args);
        };
        var callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func.apply(context, args);
    };
};

(function (window, undefined) {
    window.Asc.plugin.init = function () {
        const elements = {
            searchInput: document.getElementById('search_input'),
            search: document.getElementById('search'),
            searchResult: document.getElementById('search_result'),
            authForm: document.getElementById('auth'),
            btnRemoveKey: document.getElementById('btn_remove_key')
        }
        const handleSubmit = (e) => {
            const api_key = e.target.elements['api_key'].value;
            localStorage.setItem("x-api-key", api_key);
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'block';
        }

        const getSearchResult = async (query) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch('https://base/api/search?' + new URLSearchParams({query: query}),
                {headers: {"Accept": "application/json",  "x-api-key": key}});
            return await response.json();
        }

        const getArticleById = async (article_id) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`https://base/api/articles/${article_id}`,
                {headers: {"Accept": "application/json",  "x-api-key": key}});
            return await response.json();
        }

        const getArticleStringById = async (article_id) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`https://base/api/articles/${article_id}/str`,
                {headers: {"Accept": "application/json",  "x-api-key": key}});
            return await response.json();
        }

        const handleSearchResultItemClick = (e) => {
            const article_id = e.target.getAttribute('data-article-id')
            console.log(article_id)
            getArticleStringById(article_id)
                .then((resp) => {
                    const link = resp['link']
                    console.log(link)
                    window.Asc.plugin.executeMethod("PasteText", [link])
                })
        }

        const createListResult = (response) => {
            elements.searchResult.innerHTML = null
            const list = document.createElement('ul');
            list.setAttribute('id', 'search_result_list');
            elements.searchResult.appendChild(list)
            const renderSearchResult = (element, index, array) => {
                let articleId;
                const li = document.createElement('li');
                li.setAttribute('class','search-result-item');
                list.appendChild(li);

                if (element['_index'] === 'files') {
                    articleId = element['fields']['articles'][0]
                    li.innerHTML=li.innerHTML + element['fields']['file_name'][0]
                    li.classList.add('file')
                }
                if (element['_index'] === 'articles') {
                    articleId = element['_id'];
                    li.innerHTML=li.innerHTML + element['fields']['title'][0]
                    li.classList.add('article')
                }
                li.setAttribute('data-article-id', articleId)
                li.onclick = handleSearchResultItemClick;
                /*
                getArticleById(articleId).then((article) => {
                    const li = document.createElement('li');
                    li.setAttribute('class','item');
                    list.appendChild(li);
                    console.log(li.innerHTML);
                    li.innerHTML=li.innerHTML + article['data'][0]['title'];
                })
                 */

            }
            response['hits']['hits'].forEach(renderSearchResult);
        }

        const handleChange = (e) => {
            const query = e.target.value;
            const searchResult = getSearchResult(query)
            searchResult.then((response) => createListResult(response))
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
        elements.searchInput.onkeyup = debounce(handleChange, 300);
    };
    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);