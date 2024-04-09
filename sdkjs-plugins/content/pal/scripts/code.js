const localStorageItemsKey = {
    docId: 'pal-document-id',
    palDoc: 'pal-doc',
    apiKey: 'x-api-key'
}

const BASE_URI = 'https://base/api'

const getApiKey = () => {
    return localStorage.getItem(localStorageItemsKey.apiKey)
}

const setApiKey = (value) => {
    return localStorage.setItem(localStorageItemsKey.apiKey, value)
}

const delApiKey = () => {
    return localStorage.removeItem(localStorageItemsKey.apiKey)
}

const getDocId = () => {
    return localStorage.getItem(localStorageItemsKey.docId)
}

const setDocId = (value) => {
    return localStorage.setItem(localStorageItemsKey.docId, value)
}

const delDocId = () => {
    return localStorage.removeItem(localStorageItemsKey.docId)
}

const headers = {
    "Accept": "application/json",
    "Content-Type": "application/json",
    "x-api-key": getApiKey()
}

const saveDocumentToLocalStorage = (palDoc) => {
    return localStorage.setItem(localStorageItemsKey.palDoc, JSON.stringify(palDoc))
}

const getDocumentFromLocalStorage = () => {
    const pal_document = localStorage.getItem(localStorageItemsKey.palDoc)
    return JSON.parse(pal_document)
}

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

        const readDocId = () => {
            const docId = localStorage.getItem(localStorageItemsKey.docId)
            if (docId) return docId
            else return false
        }


        const createNewDocumentInPal = async () => {
            const key = localStorage.getItem(localStorageItemsKey.apiKey);
            const response = await fetch(`${BASE_URI}/documents`,
                {
                    method: "POST",
                    headers: headers,
                    body: JSON.stringify({articles: []})
                }
            );
            const data =  await response.json();
            if (data.code === 200) saveDocumentToLocalStorage(data.data[0])
            return data
        }
        if (!readDocId()) {
            createNewDocumentInPal()
                .then((resp) => {
                    setDocId(resp.id)
                    getDocumentById(resp.id).then((data) => {
                        if (data.code === 200) {
                            console.log(data.data[0])
                        }
                    })
                })
                .catch((err) => console.error(err))
        }

        const getDocumentById = async (document_id) => {
            const response = await fetch(`${BASE_URI}/documents/${document_id}`, {
                headers: headers
            })
            const data = await response.json()
            if (data.code === 200 && data.data.length === 1) {
                const document = data.data[0]
                saveDocumentToLocalStorage(document)
                return data.data[0]
            }
        }
        const updateDocumentById = async (document_id, articles) => {
            const response = await fetch(`${BASE_URI}/documents/${document_id}`, {
                method: 'PUT',
                headers: headers,
                body: JSON.stringify({
                    articles: articles
                })
            })
            const data = await response.json()
            if (data.code === 200) {
                saveDocumentToLocalStorage(data.data[0])
            }
            return data.data[0]
        }

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
            const response = await fetch(`${BASE_URI}/search?` + new URLSearchParams({query: query}),
                {headers: headers});
            return await response.json();
        }

        const getArticleById = async (article_id) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`${BASE_URI}/articles/${article_id}`,
                {headers: {"Accept": "application/json",  "x-api-key": key}});
            return await response.json();
        }

        const getArticleStringById = async (article_id) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`${BASE_URI}/articles/${article_id}/str`,
                {headers: {"Accept": "application/json",  "x-api-key": key}});
            return await response.json();
        }

        const handleSearchResultItemClick = async (e) => {
            const article_id = e.target.getAttribute('data-article-id')
            const {id, articles} = getDocumentFromLocalStorage();
            updateDocumentById(id, [...articles, article_id])
                .then((response) => {
                    console.log('Document updated',response)
                    Asc.scope.article_id = article_id;
                    window.Asc.plugin.callCommand(() => {
                        const oDocument = Api.GetDocument();
                        const oRange = oDocument.GetRangeBySelect();
                        //[${Asc.scope.article_id}]
                        console.log(oRange.GetText())
                        oRange.AddText(`[1] `, 'after')
                        oRange.AddRef
                        console.log(oRange.GetText())
                    })
                })

            /*
            getArticleStringById(article_id)
                .then((resp) => {
                    const link = resp['link']
                    Asc.scope.article_id = article_id;
                    Asc.scope.link = link
                    window.Asc.plugin.callCommand(() => {
                        const createBibliography = (oDocument) => {
                            const oTocPr = {
                                "ShowPageNums": true,
                                "RightAlgn": true,
                                "LeaderType": "dot",
                                "FormatAsLinks": true,
                                "BuildFrom": {"OutlineLvls": 9},
                                "TocStyle": "standard"
                            };

                            oDocument.AddTableOfContents(oTocPr)
                        }
                        const createCrossRef = (bookmarkName) => {
                            const currentSelect = oDocument.GetRangeBySelect();
                            const oParagraph = currentSelect.GetParagraph(0)
                            oParagraph.AddBookmarkCrossRef("aboveBelow", bookmarkName);
                        }
                        const addArticleAndAddBookmark = (oParagraph) => {
                            const oRun = oParagraph.AddText(Asc.scope.link);
                            const oRange = oRun.GetRange(0, 3);
                            oRange.AddBookmark(Asc.scope.article_id)
                        }
                        const oDocument = Api.GetDocument();
                        const oParagraphs = oDocument.GetAllParagraphs();
                        let isToxFind = false;
                        oParagraphs.map((oParagraph) => {
                            const oRanges = oParagraph.Search("Библиография", false)
                            if (oRanges[0]) {
                                // Найден Список литературы
                                isToxFind = true;
                                oParagraph.AddLineBreak();
                                addArticleAndAddBookmark(oParagraph)
                                createCrossRef(Asc.scope.article_id);
                                createBibliography(oDocument);
                            }
                        });
                        if (!isToxFind) {
                            const oParagraph = Api.CreateParagraph()
                            oParagraph.SetSpacingLine(240)
                            oParagraph.AddPageBreak();
                            oDocument.Push(oParagraph);
                            oParagraph.AddText("Библиография")
                            const header = oParagraph.GetRange(0, 13)
                            header.SetBold();
                            header.SetFontSize(16);
                            header.AddBookmark('bibliography');
                            oParagraph.AddLineBreak();
                            addArticleAndAddBookmark(oParagraph)
                            createCrossRef(Asc.scope.article_id);
                        }
                    }, false);
                })

             */
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