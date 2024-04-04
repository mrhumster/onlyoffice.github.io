function debounce(func, wait, immediate) {
    let timeout;
    return function() {
        const context = this, args = arguments;
        const later = function () {
            timeout = null;
            if (!immediate) func.apply(context, args);
        };
        const callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func.apply(context, args);
    };
}

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
                    Asc.scope.link = link
                    window.Asc.plugin.callCommand(() => {
                        const oDocument = Api.GetDocument();
                        const length = oDocument.GetElementsCount()
                        if (oDocument.Search("Список литературы").length === 0) {
                            const oParagraph = Api.CreateParagraph()
                            oParagraph.AddText("Список литературы");
                            oParagraph.AddLineBreak();
                            oDocument.Push(oParagraph);
                        }
                        const aSearch = oDocument.Search("Список литературы");
                        console.log(aSearch)
                        aSearch[0].AddText(Asc.scope.link);

                    }, false);
                })
        }

        const createListResult = (response) => {
            elements.searchResult.innerHTML = null
            const list = document.createElement('ul');
            list.setAttribute('id', 'search_result_list');
            elements.searchResult.appendChild(list)

            const createFileSearchResult = (element) => {
                const articleId = element['fields']['articles'][0];
                const container = document.createElement('div')
                container.style.display = 'flex';
                container.style.flexDirection = 'column';
                const fileName = document.createElement('div');
                fileName.style.display = 'flex';
                fileName.appendChild(document.createTextNode(element['fields']['file_name'][0]));
                const btnContainer = document.createElement('div');
                btnContainer.style.display = 'flex';
                const addButton = document.createElement('button')
                addButton.appendChild(document.createTextNode('Вставить в документ'))
                addButton.classList.add('btn-text-default');
                addButton.setAttribute('data-article-id', articleId)
                addButton.onclick = handleSearchResultItemClick;
                btnContainer.appendChild(addButton);
                container.appendChild(fileName);
                container.appendChild(btnContainer);
                return container;
            }

            const renderSearchResult = (element, index, array) => {
                let articleId;
                const li = document.createElement('li');
                li.setAttribute('class','search-result-item');
                list.appendChild(li);

                if (element['_index'] === 'files') {
                    li.appendChild(createFileSearchResult(element));
                    li.classList.add('file')
                }
                if (element['_index'] === 'articles') {
                    articleId = element['_id'];
                    li.innerHTML=li.innerHTML + element['fields']['title'][0]
                    li.classList.add('article')
                }
            }
            response['hits']['hits'].forEach(renderSearchResult);
        }

        const handleChange = (e) => {
            const query = e.target.value;
            const searchResult = getSearchResult(query)
            searchResult
                .then((response) => createListResult(response))
                .catch((error) => {
                    handleRemoveKey()
                })
        }

        const handleRemoveKey = () => {
            localStorage.removeItem("x-api-key")
            elements.authForm.style.display = 'block';
            elements.search.style.display = 'none';
        }

        const key = localStorage.getItem("x-api-key");
        if (key) {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'flex';
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