const localStorageItemsKey = {
    docId: 'pal-document-id',
    palDoc: 'pal-doc',
    apiKey: 'x-api-key',
    articlesCites: 'pal-articles-cites',
    bibliographyContentControl: 'pal-bib-internal-id',
    styles: 'pal-styles',
    locale: 'pal-locale'
}

const BASE_URI = 'https://pal.test.vniigaz.local/api'

const truncateString = (string, length = 20) => {
    if (string.length > length)
        return `${string.slice(0, Math.trunc(length / 2))}…${string.slice(string.length - Math.trunc(length / 2))}`
    return string
}

const erasePalArtifactsInLocalStorage = () => {
    localStorage.removeItem(localStorageItemsKey.docId)
    localStorage.removeItem(localStorageItemsKey.palDoc)
    localStorage.removeItem(localStorageItemsKey.articlesCites)
    localStorage.removeItem(localStorageItemsKey.bibliographyContentControl)
}

const urlify = (text) => {
    const urlRegex = /(https?:\/\/[^\s]+)/g;
    return text.replace(urlRegex, function(url) {
        return '<a href="' + url + '" target="_blank">' + url + '</a>';
    })
}

const normalizeHighLight = (s) => {
    return urlify(s.replace('-\n', '').replace('- \n', ''));
}

const getApiKey = () => {
    return localStorage.getItem(localStorageItemsKey.apiKey)
}

const setApiKey = (value) => {
    return localStorage.setItem(localStorageItemsKey.apiKey, value)
}

const delApiKey = () => {
    return localStorage.removeItem(localStorageItemsKey.apiKey)
}

const setStyles = (value) => {
    return localStorage.setItem(localStorageItemsKey.styles, JSON.stringify(value));
}

const getStyles = () => {
    return JSON.parse(localStorage.getItem(localStorageItemsKey.styles));
}

const setLocale = (value) => {
    return localStorage.setItem(localStorageItemsKey.locale, JSON.stringify(value));
}

const getLocale = () => {
    return JSON.parse(localStorage.getItem(localStorageItemsKey.locale));
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

const getArticleStringsFromLocalStorage = () => {
    const cites = localStorage.getItem(localStorageItemsKey.articlesCites)
    return JSON.parse(cites)
}


const saveArticleStringsToLocalStorageAndReturnDifference = (lst) => {
    const oldArticleList = getArticleStringsFromLocalStorage();
    const indexDifferences = {}
    if (oldArticleList) {
        for (let article_id in oldArticleList) {
            indexDifferences[article_id] = {
                "was": oldArticleList.hasOwnProperty(article_id) ? oldArticleList[article_id].index : null,
                "now": lst.hasOwnProperty(article_id) ? lst[article_id].index : null
            }
        }
        for (let article_id in lst) {
            if (!indexDifferences.hasOwnProperty(article_id))
                indexDifferences[article_id] = {
                    "was": oldArticleList.hasOwnProperty(article_id) ? oldArticleList[article_id].index : null,
                    "now": lst.hasOwnProperty(article_id) ? lst[article_id].index : null
                }
        }
    }
    localStorage.setItem(localStorageItemsKey.articlesCites, JSON.stringify(lst));
    return indexDifferences
}

const saveContentControlBibliographyInternalIdToLocalStorage = (value) => {
    return localStorage.setItem(localStorageItemsKey.bibliographyContentControl, value)
}

const getContentControlBibliographyInternalIdFromLocalStorage = () => {
    /* Вернуть идентификатор блока ContentControl для библиографии */
    return localStorage.getItem(localStorageItemsKey.bibliographyContentControl)
}

function debounce(func, wait, immediate) {
    let timeout;
    return function() {
        const context = this, args = arguments;
        const later = function() {
            timeout = null;
            if (!immediate) func.apply(context, args);
        };
        const callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func.apply(context, args);
    };
}


(function(window, undefined) {
    window.Asc.plugin.init = function() {
        erasePalArtifactsInLocalStorage();
        const elements = {
            searchInput: document.getElementById('search_input'),
            search: document.getElementById('search'),
            searchResult: document.getElementById('search_result'),
            authForm: document.getElementById('auth'),
            btnRemoveKey: document.getElementById('btn_remove_key'),
            btnArticleList: document.getElementById('btn_article_list'),
            btnSearch: document.getElementById('btn_search'),
            articleList: document.getElementById('articles_list'),
            errorList: document.getElementById('error_list'),

        }

        const showError = (errorText) => {
            elements.errorList.innerHTML = null;
            elements.errorList.style.display = 'flex'
            const closeButton = document.createElement('button')
            closeButton.appendChild(document.createTextNode('Закрыть'));
            closeButton.classList.add('btn-text-default');
            closeButton.classList.add('btn-close-error');
            const message = document.createElement('div')
            message.classList.add('text-error')
            message.appendChild(document.createTextNode(errorText))
            elements.errorList.appendChild(message)
            elements.errorList.appendChild(closeButton)
            closeButton.addEventListener('click', () => {
                elements.errorList.style.display = 'none';
            })
        }

        const createNewDocumentInPal = async () => {
            const response = await fetch(`${BASE_URI}/documents`,
                {
                    method: "POST",
                    headers: headers,
                    body: JSON.stringify({ articles: [] })
                }
            );
            const data = await response.json();
            saveDocumentToLocalStorage(data)
            setDocId(data.id)
            return data
        }

        const fetchStyles = async () => {
            // TODO: Переделать на получение стандартных стилей пользователя
            const response = await fetch(`${BASE_URI}/articles/styles/`, { headers: headers })
            const data = await response.json();
            data.then((value) => {
                setStyles(value);
            })
            return data
        }

        const getDocumentById = async (document_id) => {
            const response = await fetch(`${BASE_URI}/documents/${document_id}`, {
                headers: headers
            })
            const data = await response.json()
            saveDocumentToLocalStorage(data)
            return data
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
            saveDocumentToLocalStorage(data)
            return data
        }

        const getBibliographyByDocumentId = async (fmt, locale) => {
            const document = getDocumentFromLocalStorage();
            const response = await fetch(`${BASE_URI}/documents/${document.id}/str?${new URLSearchParams({
                fmt: fmt,
                locale: locale
            })}`,
                { headers: headers });
            if (response.status === 400) {
                const error = await response.json();
                throw new Error(error['detail'])
            }
            return await response.json()
        }

        const getCurrentCiteSettings = async () => {
            const response = await fetch(`${BASE_URI}/users`, { headers: headers });
            return await response.json();
        }

        const onClickSetApiKeyHandle = (e) => {
            const api_key = e.target.elements['api_key'].value;
            localStorage.setItem("x-api-key", api_key);
            elements.authForm.style.display = 'none';
            elements.articleList.style.display = 'none';
            elements.search.style.display = 'none';
        }

        const onClickRemoveArticleFromDocumentHandler = async (e) => {
            const article_id = e.target.getAttribute('data-article-id');
            const articles = getDocumentFromLocalStorage().articles.filter((item) => item !== article_id);
            const document_id = getDocumentFromLocalStorage().id;
            await updateDocumentById(document_id, articles)
                .then((response) => {
                    updateArticleList()
                        .then(() => console.log('Bibliography updated in backend'));
                    updateBibliographyInDocument()
                        .then(() => console.log('Bibliography updated in document'));
                })
        }

        const onClickPastInTextButtonHandler = async (e) => {
            const article_id = e.target.getAttribute('data-article-id');
            const index = getArticleStringsFromLocalStorage()[article_id]['index']
            Asc.scope.bookmark_id = article_id
            Asc.scope.bookmark_idx = index

            this.executeMethod('AddContentControl', [2, {
                "Id": Asc.scope.bookmark_idx,
                "Tag": Asc.scope.bookmark_id,
                "Lock": 3,
                "Appearance": 2
            }], (_cc) => {
                const arrDocuments = [{
                    "Props": {
                        "InternalId": _cc.InternalId,
                        "Inline": true
                    },
                    "Script": `
                        const cite = Api.CreateParagraph();
                        cite.AddText("[${_cc.Id}]");
                        Api.GetDocument().InsertContent([cite], true, {KeepTextOnly: true});
                    `
                }]
                this.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
            });
        }

        const updateReferenceInDocument = (diff) => {
            const diffWithoutNotChanged = {}
            for (let key in diff) {
                if (diff[key].was !== diff[key].now) {
                    diffWithoutNotChanged[key] = diff[key]
                }
            }
            Asc.scope.diffWithoutNotChanged = diffWithoutNotChanged;
            this.executeMethod("GetAllContentControls", null, (data) => {
                for (let i = 0; i < data.length; i++) {
                    if (data[i].Tag.indexOf('bibliography') === -1) {
                        if (Asc.scope.diffWithoutNotChanged.hasOwnProperty(data[i].Tag)) {
                            const citeId = Asc.scope.diffWithoutNotChanged[data[i].Tag].now
                            const text = citeId ? `[${citeId}]` : ''
                            const arrDocuments = [{
                                "Props": {
                                    "InternalId": data[i].InternalId,
                                    "Inline": true
                                },
                                "Script": `
                                    const cite = Api.CreateParagraph();
                                    cite.AddText("${text}");
                                    Api.GetDocument().InsertContent([cite], true, {KeepTextOnly: true});
                                `
                            }]
                            this.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
                            if (!Asc.scope.diffWithoutNotChanged[data[i].Tag].now) {
                                this.executeMethod("RemoveContentControl", [data[i].InternalId]);
                            }
                        }
                    }
                }
            })
        }

        const onAddContentControl = (_cc) => {
            saveContentControlBibliographyInternalIdToLocalStorage(_cc.InternalId)
            const arrDocuments = [{
                "Props": {
                    "InternalId": _cc.InternalId,
                    "Id": getDocId(),
                    "Appearance": 2
                },
                "Script": `
                        const cites = JSON.parse(localStorage.getItem('${localStorageItemsKey.articlesCites}'));
                        const oDocument = Api.GetDocument();
                        const articlesStrings = cites['articles'];
                        const bg = Api.CreateParagraph();
                        oDocument.InsertContent([bg]);
                        const title = Api.CreateParagraph();
                        title.SetJc('center');
                        const oTextPr = oDocument.GetDefaultTextPr();
                        oTextPr.SetBold(true);
                        title.AddText('Библиография');
                        bg.InsertParagraph(title, 'before', true);
                        oTextPr.SetBold(false);
                        const keys = Object.keys(cites);
                        keys.map((key) => {
                           const cite = cites[key];
                           const end = cite.index.toString().length;
                           const oParagraph = Api.CreateParagraph();
                           oParagraph.AddText(cite.refer);
                           bg.InsertParagraph(oParagraph, 'before', true);
                           const oRange = oParagraph.GetRange(0,end);
                           oRange.AddBookmark(key);
                        });
                        `
            }]
            this.executeMethod("InsertAndReplaceContentControls", [arrDocuments], (_re) => {
                console.log(_re)
            });
        }

        const addBibBtnClickHandler = async () => {
            /*
             * Добавление библиографии в документ
             */
            const style = getStyles();
            const locale = getLocale();
            await getBibliographyByDocumentId(style.defaultCiteStyleId, locale.id)
                .then((data) => {
                    const diff = saveArticleStringsToLocalStorageAndReturnDifference(data.bibliography)
                    updateReferenceInDocument(diff);
                    const documentId = getDocId();

                    const onMoveCursorToEnd = () => {
                        this.executeMethod('AddContentControl', [1, {
                            "Id": documentId,
                            "Tag": `bibliography-${documentId}`,
                            "Lock": 3,
                            "Alias": "Библиография",
                            "Appearance": 2
                        }], onAddContentControl);
                    }
                    this.callCommand(() => {
                        const oDocument = Api.GetDocument();
                        const oParagraph = oDocument.GetElement(oDocument.GetElementsCount() - 1)
                        oParagraph.AddPageBreak();
                    })
                    this.executeMethod("MoveCursorToEnd", [true], onMoveCursorToEnd);
                })
                .catch((error) => showError(error))
        }

        const updateArticleList = async () => {
            const doc = getDocumentFromLocalStorage();

            if (!doc) throw 'Document not created'

            if (doc) {
                elements.articleList.innerHTML = null;
                /*
                const title = document.createElement('h2');
                title.style.gridColumn = "span 3 / span 3";
                title.style.textAlign = "left";
                title.appendChild(document.createTextNode('Библиография'));
                 */
                const btnCont = document.createElement('div');
                btnCont.classList.add('buttons_container');
                btnCont.style.gridColumn = "span 3 / span 3";
                const addBibBtn = document.createElement('button');
                addBibBtn.appendChild(document.createTextNode('Вставить библиографию в документ'));
                addBibBtn.classList.add("w-xxl", "border", "hover-hl");
                addBibBtn.id = 'btn_add_bib';
                addBibBtn.onclick = addBibBtnClickHandler;
                btnCont.appendChild(addBibBtn)
                // elements.articleList.appendChild(title)
                elements.articleList.appendChild(btnCont)
                doc['articles'].map((articleId) => {
                    const response = getArticleById(articleId);
                    response
                        .then((article) => {
                            const articleTitle = document.createElement('div')
                            articleTitle.classList.add('article-title')
                            articleTitle.style.display = 'flex';
                            articleTitle.style.justifyContent = 'center';
                            articleTitle.style.alignContent = 'center';
                            articleTitle.style.flexDirection = 'column';
                            articleTitle.setAttribute('data-article-id', article['id'])
                            articleTitle.appendChild(document.createTextNode(article['title']))

                            const pastInTextButton = document.createElement('button');
                            pastInTextButton.classList.add("w-xxs", "no-border", "hover-hl", "add-article-button");
                            pastInTextButton.setAttribute('data-article-id', article['id']);

                            pastInTextButton.addEventListener('click', onClickPastInTextButtonHandler)
                            pastInTextButton.setAttribute('data-title-left', 'Вставить в документ ссылку');

                            const removeBtn = document.createElement('button');
                            removeBtn.classList.add("w-xxs", "no-border", "hover-hl", "remove-article-button");
                            removeBtn.setAttribute('data-article-id', article['id']);
                            removeBtn.addEventListener('click', onClickRemoveArticleFromDocumentHandler);
                            removeBtn.setAttribute('data-title-left', 'Удалить из библиографии');

                            elements.articleList.appendChild(articleTitle);
                            elements.articleList.appendChild(pastInTextButton);
                            elements.articleList.appendChild(removeBtn);
                        })
                        .catch((error) => showError(error))
                });
            }
        }

        const showList = async () => {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'grid';
            elements.btnArticleList.classList.add('active');
            elements.btnSearch.classList.remove('active');

        }
        const showSearch = () => {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'flex';
            elements.btnSearch.classList.add('active');
            elements.btnArticleList.classList.remove('active');
            elements.articleList.style.display = 'none';
        }

        const showSettings = () => {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'none';
        }

        const getSearchResult = async (query) => {
            const response = await fetch(`${BASE_URI}/search?` + new URLSearchParams({ query: query }),
                { headers: headers });
            return await response.json();
        }

        const getArticleById = async (article_id) => {
            const response = await fetch(`${BASE_URI}/articles/${article_id}`,
                { headers: headers });
            return await response.json();
        }

        const updateBibliographyInDocument = async () => {
            const contentControlInternalId = getContentControlBibliographyInternalIdFromLocalStorage();
            if (!contentControlInternalId) throw 'Not found internal id for bibliography content control in local storage';
            const documentId = getDocId();
            if (!documentId) throw 'Not found document id in local storage'
            const style = getStyles();
            const locale = getLocale();
            await getBibliographyByDocumentId(style.defaultCiteStyleId, locale.id)
                .then((data) => {
                    const diff = saveArticleStringsToLocalStorageAndReturnDifference(data.bibliography)
                    updateReferenceInDocument(diff);
                    const arrDocuments = [{
                        "Props": {
                            "InternalId": contentControlInternalId,
                            "Id": documentId,
                            "Appearance": 2
                        },
                        "Script": `
                        const cites = JSON.parse(localStorage.getItem('${localStorageItemsKey.articlesCites}'));
                        const oDocument = Api.GetDocument();
                        const articlesStrings = cites['articles'];
                        const bg = Api.CreateParagraph();
                        oDocument.InsertContent([bg]);
                        const title = Api.CreateParagraph();
                        title.SetJc('center');
                        const oTextPr = oDocument.GetDefaultTextPr();
                        oTextPr.SetBold(true);
                        title.AddText('Библиография');
                        bg.InsertParagraph(title, 'before', true);
                        oTextPr.SetBold(false);
                        const keys = Object.keys(cites);
                        keys.map((key) => {
                           const cite = cites[key];
                           const end = cite.index.toString().length;
                           const oParagraph = Api.CreateParagraph();
                           oParagraph.AddText(cite.refer);
                           bg.InsertParagraph(oParagraph, 'before', true);
                           const oRange = oParagraph.GetRange(0,end);
                           oRange.AddBookmark(key);
                        });
                        `
                    }]
                    this.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
                })
                .catch((error) => showError(error))

        }

        const onClickSearchResultItemHandler = async (e) => {
            const article_id = e.target.getAttribute('data-article-id')
            const { id, articles } = getDocumentFromLocalStorage();
            updateDocumentById(id, [...articles, article_id])
                .then((response) => {
                    updateArticleList()
                        .then(() => console.log('Bibliography updated in backend'));
                    updateBibliographyInDocument()
                        .then(() => console.log('Bibliography updated in document'));
                })
        }

        const createListResult = (response) => {
            elements.searchResult.innerHTML = null

            const renderSearchResult = async (element, index, array) => {
                // const articleSearchItem = await createArticleSearchItem(articleId)
                const container = document.createElement('div')
                container.classList.add('search-item-container')
                const titleCont = document.createElement('div');
                titleCont.style.display = 'flex';
                titleCont.style.flexDirection = 'column';
                if (element.title) {
                    const title = document.createElement('div')
                    title.classList.add('search-item-container-title')
                    title.appendChild(document.createTextNode(element.title))
                    titleCont.appendChild(title)
                }

                const buttonsCont = document.createElement('div')
                buttonsCont.classList.add('search-item-container-buttons-container')
                const addButton = document.createElement('button')
                const openInWebButton = document.createElement('button')
                openInWebButton.classList.add("w-xxs", "no-border", "hover-hl", "open-in-web-btn");
                openInWebButton.setAttribute('data-title', 'Открыть ссылку в web версии')
                addButton.classList.add("w-xxs", "no-border", "hover-hl", "add-to-document-btn");
                addButton.setAttribute('data-article-id', element.id)
                addButton.setAttribute('data-title', 'Добавить ссылку в библиографию')
                addButton.onclick = onClickSearchResultItemHandler;
                buttonsCont.appendChild(addButton);
                buttonsCont.appendChild(openInWebButton);
                titleCont.appendChild(buttonsCont)
                if (element.file_name) {
                    const fileName = document.createElement('div')
                    fileName.classList.add('search-item-container-file')
                    fileName.setAttribute('data-title-center', element.file_name);
                    fileName.appendChild(document.createTextNode(truncateString(element.file_name, 50)))
                    titleCont.appendChild(fileName)
                }
                if (element.highlights) {
                    for (let [name, value] of Object.entries(element.highlights)) {
                        value.map((_hl) => {
                            const hl = document.createElement('div');
                            hl.insertAdjacentHTML('beforeend', normalizeHighLight(_hl));
                            hl.classList.add('search-item-container-highlight');
                            titleCont.appendChild(hl);
                        })
                    }
                }


                container.appendChild(titleCont);


                elements.searchResult.appendChild(container);

            }
            response.forEach(renderSearchResult);
        }

        const handleChange = (e) => {
            const query = e.target.value;
            const searchResult = getSearchResult(query)
            searchResult
                .then((response) => {
                    createListResult(response)
                })
                .catch((error) => {
                    handleRemoveKey()
                })
        }

        const handleRemoveKey = () => {
            localStorage.removeItem("x-api-key")
            elements.authForm.style.display = 'flex';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'none';
            elements.btnRemoveKey.style.display = 'none'
            elements.btnSearch.style.display = 'none';
            elements.btnArticleList.style.display = 'none';
        }

        const searchBibliographyInDocument = () => {
            this.executeMethod("GetAllContentControls", null, (data) => {
                let founded = false;
                for (let i = 0; i < data.length; i++) {
                    if (founded) break
                    if (data[i].Tag.indexOf('bibliography') !== -1) {
                        founded = true;
                        setDocId(data[i].Tag.split('-')[1]);
                        saveContentControlBibliographyInternalIdToLocalStorage(data[i].InternalId);
                        let documentId = getDocId();
                        getDocumentById(documentId)
                            .then((data) => {
                                console.log('Document received from backend and save in local storage')
                                updateArticleList().then(() => console.log('Bibliography in plugin updated'))
                                const style = getStyles();
                                const locale = getLocale();
                                getBibliographyByDocumentId(style.defaultCiteStyleId, locale.id)
                                    .then((data) => saveArticleStringsToLocalStorageAndReturnDifference(data.bibliography))

                            })
                            .catch((err) => console.error('Document not received from backend', err))
                    }
                }
                if (!founded) {
                    createNewDocumentInPal()
                        .then((data) => {
                            console.info('New document created on backend', data)
                        })
                        .catch((error) => console.error(error))
                }
            });
        }

        getCurrentCiteSettings()
            .then(data => {
                console.log('Getting curent style');
                setStyles(data.style);
                setLocale(data.locale);
            })
            .catch(error => console.log('Problem with getting settings.', error));

        searchBibliographyInDocument();

        const key = getApiKey();

        if (key) {
            elements.authForm.style.display = 'none';
            elements.btnRemoveKey.style.display = 'block';
            elements.articleList.style.display = 'none'
            elements.search.style.display = 'flex';
            elements.btnSearch.classList.add('active')
            elements.btnArticleList.style.display = 'block';
            elements.btnSearch.style.display = 'block';
        } else {
            elements.authForm.style.display = 'block';
            elements.btnRemoveKey.style.display = 'none';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'none';
            elements.btnArticleList.style.display = 'none';
            elements.btnSearch.style.display = 'none';
        }

        elements.authForm.onsubmit = onClickSetApiKeyHandle;
        elements.btnRemoveKey.onclick = handleRemoveKey;
        elements.searchInput.onkeyup = debounce(handleChange, 300);
        elements.btnArticleList.onclick = showList;
        elements.btnSearch.onclick = showSearch;

    };
    window.Asc.plugin.button = (id) => {
        erasePalArtifactsInLocalStorage();
        window.Asc.plugin.executeCommand("close", "");
    };

})(window, undefined);
