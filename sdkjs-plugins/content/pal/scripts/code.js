const localStorageItemsKey = {
    docId: 'pal-document-id',
    palDoc: 'pal-doc',
    apiKey: 'x-api-key',
    articlesCites: 'pal-articles-cites'
}

const BASE_URI = 'https://base/api'

const truncateString = (string, length = 20) => {
    if (string.length > length)
        return `${string.slice(0, Math.trunc(length / 2))}…${string.slice(string.length - Math.trunc(length / 2))}`
    return string
}


const urlify = (text) => {
    const urlRegex = /(https?:\/\/[^\s]+)/g;
    return text.replace(urlRegex, function (url) {
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

const saveArticleStringsToLocalStorage = (lst) => {
    return localStorage.setItem(localStorageItemsKey.articlesCites, JSON.stringify(lst))
}

const getArticleStringsFromLocalStorage = () => {
    const cites = localStorage.getItem(localStorageItemsKey.articlesCites)
    return JSON.parse(cites)
}

function debounce(func, wait, immediate) {
    let timeout;
    return function () {
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

        const docId = readDocId();

        const createNewDocumentInPal = async () => {
            const key = localStorage.getItem(localStorageItemsKey.apiKey);
            const response = await fetch(`${BASE_URI}/documents`,
                {
                    method: "POST",
                    headers: headers,
                    body: JSON.stringify({articles: []})
                }
            );
            const data = await response.json();
            saveDocumentToLocalStorage(data)
            setDocId(data.id)
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

        const getBiblioByDocumentId = async (fmt, locale) => {
            const document = getDocumentFromLocalStorage();
            const response = await fetch(`${BASE_URI}/documents/${docId}/str?${new URLSearchParams({
                    fmt: fmt,
                    locale: locale
                })}`,
                {headers: headers});

            const data = await response.json();
            saveArticleStringsToLocalStorage(data.bibliography)
            return data
        }

        const elements = {
            searchInput: document.getElementById('search_input'),
            search: document.getElementById('search'),
            searchResult: document.getElementById('search_result'),
            authForm: document.getElementById('auth'),
            btnRemoveKey: document.getElementById('btn_remove_key'),
            btnArticleList: document.getElementById('btn_article_list'),
            btnSearch: document.getElementById('btn_search'),
            articleList: document.getElementById('articles_list')
        }
        const handleSubmit = (e) => {
            const api_key = e.target.elements['api_key'].value;
            localStorage.setItem("x-api-key", api_key);
            elements.authForm.style.display = 'none';
            elements.articleList.style.display = 'none';
            elements.search.style.display = 'none';
        }

        const removeArticleFromList = async (article_id) => {
            const articles_list = document.querySelectorAll(`[data-article-id='${article_id}']`)
            const article_array = [...articles_list]
            article_array.forEach(div => div.remove())
        }

        const removeArticleFromDocumentClickHandler = async (e) => {
            // TODO: Удаление ссылок из документов
            const article_id = e.target.getAttribute('data-article-id');
            const articles = getDocumentFromLocalStorage().articles.filter((item) => item !== article_id);
            const document_id = getDocumentFromLocalStorage().id;
            const updated_doc = await updateDocumentById(document_id, articles);
            await removeArticleFromList(article_id);
        }

        const onAddCiteCC = async (_cc) => {
            localStorage.setItem('_cc', JSON.stringify(_cc));
            console.log(_cc)
            const arrDocuments = [{
                "Props": {
                    "InternalId": _cc.InternalId,
                    "Inline": true
                },
                "Script":
                    "const _cc = JSON.parse(localStorage.getItem('_cc'));" +
                    "const index = _cc.Tag;" +
                    "const oDocument = Api.GetDocument();" +
                    "let cite = Api.CreateParagraph();" +
                    "cite.AddText('ddd');" +
                    "oDocument.InsertContent([cite]);"

            }]
            this.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
            this.executeMethod("PasteHtml", [`<span id="${_cc.Tag}">${_cc.Tag}</span>`])
        }

        const pastInTextButtonClickHandler = async (e) => {
            const article_id = e.target.getAttribute('data-article-id');
            const index = getArticleStringsFromLocalStorage()[article_id]['index']
            Asc.scope._article_id = article_id;
            Asc.scope._index = index
            // this.executeMethod('AddContentControl', [2, {"Id": article_id, "Tag": index, "Lock": 3}], onAddCiteCC);
            //this.executeMethod("PasteHtml", [`<span id=${article_id}>[${index}]</span>`]);
            this.callCommand(function () {
                    let oDocument = Api.GetDocument();
                    let oParagraph = Api.CreateParagraph();
                    oParagraph.AddText(`[${Asc.scope._index}]`);
                    oParagraph.AddBookmarkCrossRef("aboveBelow", Asc.scope._article_id);
                    oDocument.InsertContent([oParagraph]);
                },

                false);
        }

        const addBibBtnClickHandler = async () => {
            await getBiblioByDocumentId('gost-r-7-0-5-2008', 'ru-RU')
            const documentId = getDocId();
            const onAddContentControl = (_cc) => {
                localStorage.setItem('pal-bibliography-id', _cc.InternalId)
                const arrDocuments = [{
                    "Props": {
                        "InternalId": _cc.InternalId,
                        "Appearance": 2
                    },
                    "Script":
                        "const cites = JSON.parse(localStorage.getItem('pal-articles-cites'));" +
                        "const oDocument = Api.GetDocument();" +
                        "const articlesStrings = cites['articles'];" +
                        "const bg = Api.CreateParagraph();" +
                        "oDocument.InsertContent([bg]);" +
                        "const title = Api.CreateParagraph();" +
                        "title.SetJc('center');" +
                        "const oTextPr = oDocument.GetDefaultTextPr();" +
                        "oTextPr.SetBold(true);" +
                        "title.AddText('Библиография');" +
                        "bg.InsertParagraph(title, 'before', true);" +
                        "oTextPr.SetBold(false);" +
                        "const keys = Object.keys(cites);" +
                        "keys.map((key) => {" +
                        "   const cite = cites[key];" +
                        "   const oParagraph = Api.CreateParagraph();" +
                        "   oParagraph.AddText(cite.refer); " +
                        "   bg.InsertParagraph(oParagraph, 'before', true);" +
                        "   const oRange = oParagraph.GetRange(0,3);" +
                        "   oRange.AddBookmark(key);" +
                        "});"
                }]
                this.executeMethod("InsertAndReplaceContentControls", [arrDocuments], (_re) => {console.log(_re)});
            }
            const onMoveCursorToEnd = () => {
                this.executeMethod('AddContentControl', [1, {
                    "Id": documentId,
                    "Tag": "bibliography",
                    "Lock": 3,
                    "Alias": "Библиография",
                    "Appearance": 2
                }], onAddContentControl);
            }
            this.callCommand(() => {
                const oDocument = Api.GetDocument();
                const oParagpraph = oDocument.GetElement(oDocument.GetElementsCount() - 1)
                oParagpraph.AddPageBreak();
            })
            this.executeMethod("MoveCursorToEnd", [true], onMoveCursorToEnd);


        }

        const updateArticleList = async () => {
            const doc = getDocumentFromLocalStorage();
            if (doc) {
                elements.articleList.innerHTML = null;
                const title = document.createElement('h2');
                title.style.gridColumn = "span 3 / span 3";
                title.style.textAlign = "center";
                title.appendChild(document.createTextNode('Библиография'));
                const btnCont = document.createElement('div');
                btnCont.classList.add('buttons_container');
                btnCont.style.gridColumn = "span 3 / span 3";
                const addBibBtn = document.createElement('button');
                addBibBtn.appendChild(document.createTextNode('Вставить библиографию'));
                addBibBtn.classList.add('btn-text-default');
                addBibBtn.id = 'btn_add_bib';
                addBibBtn.onclick = addBibBtnClickHandler;
                btnCont.appendChild(addBibBtn)
                elements.articleList.appendChild(title)
                elements.articleList.appendChild(btnCont)
                doc['articles'].map((articleId) => {
                    const response = getArticleById(articleId);
                    response
                        .then((article) => {
                            const articleTitle = document.createElement('div')
                            articleTitle.style.display = 'flex';
                            articleTitle.style.justifyContent = 'center';
                            articleTitle.style.alignContent = 'center';
                            articleTitle.style.flexDirection = 'column';
                            articleTitle.setAttribute('data-article-id', article['id'])
                            articleTitle.appendChild(document.createTextNode(article['title']))

                            const pastInTextButton = document.createElement('button');
                            pastInTextButton.classList.add("btn-text-default");
                            pastInTextButton.setAttribute('data-article-id', article['id']);
                            pastInTextButton.appendChild(document.createTextNode('⬆'));
                            pastInTextButton.addEventListener('click', pastInTextButtonClickHandler)
                            pastInTextButton.title = 'Вставить в документ ссылку';

                            const removeBtn = document.createElement('button');
                            removeBtn.classList.add("btn-text-default");
                            removeBtn.setAttribute('data-article-id', article['id']);
                            removeBtn.appendChild(document.createTextNode('🚫'));
                            removeBtn.addEventListener('click', removeArticleFromDocumentClickHandler);
                            removeBtn.title = 'Удалить из библиографии';

                            elements.articleList.appendChild(articleTitle);
                            elements.articleList.appendChild(pastInTextButton);
                            elements.articleList.appendChild(removeBtn);
                        })
                        .catch((error) => {
                            console.log(error);
                        })
                });
            }
        }

        const showList = async () => {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'grid';
        }
        const showSearch = () => {
            elements.authForm.style.display = 'none';
            elements.search.style.display = 'flex';
            elements.articleList.style.display = 'none';
        }

        const getSearchResult = async (query) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`${BASE_URI}/search?` + new URLSearchParams({query: query}),
                {headers: headers});
            return await response.json();
        }

        const getArticleById = async (article_id) => {
            // const key = localStorage.getItem("x-api-key");
            const response = await fetch(`${BASE_URI}/articles/${article_id}`,
                {headers: headers});
            return await response.json();
        }

        const getArticleStringById = async (article_id) => {
            const key = localStorage.getItem("x-api-key");
            const response = await fetch(`${BASE_URI}/articles/${article_id}/str`,
                {headers: headers});
            return await response.json();
        }

        const handleSearchResultItemClick = async (e) => {
            const article_id = e.target.getAttribute('data-article-id')
            const {id, articles} = getDocumentFromLocalStorage();
            updateDocumentById(id, [...articles, article_id])
                .then((response) => {
                    updateArticleList().then(() => console.log('Updated'));
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

            const createArticleSearchItem = async (articleId) => {
                // TODO: Create article search item
                const article = await getArticleById(articleId)
                const container = document.createElement('div')
                container.style.display = 'grid';
                container.style.gridTemplateColumns = '85% 15%'
                const titleCont = document.createElement('div');
                titleCont.appendChild(document.createTextNode(article.title));

                const addButton = document.createElement('button')
                addButton.appendChild(document.createTextNode('➕'))
                addButton.classList.add('btn-text-default');
                addButton.setAttribute('data-article-id', articleId)
                addButton.onclick = handleSearchResultItemClick;
                container.appendChild(titleCont);
                container.appendChild(addButton);
                return container;
            }

            const renderSearchResult = async (element, index, array) => {
                // const articleSearchItem = await createArticleSearchItem(articleId)
                const container = document.createElement('div')
                container.classList.add('search_item_container')
                const titleCont = document.createElement('div');
                titleCont.style.display = 'flex';
                titleCont.style.flexDirection = 'column';
                if (element.title) {
                    const title = document.createElement('div')
                    title.classList.add('search_item_container_title')
                    title.appendChild(document.createTextNode(element.title))
                    titleCont.appendChild(title)
                }
                if (element.file_name) {
                    const fileName = document.createElement('div')
                    fileName.classList.add('search_item_container_file')
                    fileName.title = element.file_name;
                    fileName.appendChild(document.createTextNode(truncateString(element.file_name, 50)))
                    titleCont.appendChild(fileName)
                }
                if (element.highlights) {
                    for (let [name, value] of Object.entries(element.highlights)) {
                        value.map((_hl) => {
                            const hl = document.createElement('div');
                            hl.insertAdjacentHTML('beforeend', normalizeHighLight(_hl));
                            hl.classList.add('search_item_container_highlight');
                            titleCont.appendChild(hl);
                        })
                    }
                }
                const buttonsCont = document.createElement('div')
                buttonsCont.classList.add('search_item_container_buttons_container')
                const addButton = document.createElement('button')
                addButton.appendChild(document.createTextNode('Добавить в библиографию'))
                const openInWebButton = document.createElement('button')
                openInWebButton.appendChild(document.createTextNode('Открыть в браузере'))
                openInWebButton.classList.add('btn-text-default');
                openInWebButton.classList.add('open_in_web_btn');
                addButton.classList.add('btn-text-default');
                addButton.classList.add('add_to_document_btn');
                addButton.setAttribute('data-article-id', element.id)
                addButton.onclick = handleSearchResultItemClick;
                buttonsCont.appendChild(addButton);
                buttonsCont.appendChild(openInWebButton);
                container.appendChild(titleCont);
                container.appendChild(buttonsCont)

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
            elements.authForm.style.display = 'block';
            elements.search.style.display = 'none';
            elements.articleList.style.display = 'none';
            elements.btnRemoveKey.style.display = 'none'
            elements.btnSearch.style.display = 'none';
            elements.btnArticleList.style.display = 'none';
        }

        if (!getDocId()) {
            createNewDocumentInPal()
                .then((data) => console.info(data))
                .catch((error) => console.error(error))

        } else {
            getDocumentById(docId)
                .then((data) => {
                    console.log('Документ сохранен в локальное хранилище')
                })
                .catch((err) => console.error('С получение нового документа что-то пошло не так!', err))
        }

        const key = localStorage.getItem("x-api-key");

        if (key) {
            elements.authForm.style.display = 'none';
            elements.btnRemoveKey.style.display = 'block';
            elements.articleList.style.display = 'none'
            elements.search.style.display = 'flex';
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
        updateArticleList().then(() => console.log('Библиография обновлена'))
        elements.authForm.onsubmit = handleSubmit;
        elements.btnRemoveKey.onclick = handleRemoveKey;
        elements.searchInput.onkeyup = debounce(handleChange, 300);
        elements.btnArticleList.onclick = showList;
        elements.btnSearch.onclick = showSearch;
    };
    window.Asc.plugin.button = (id) => {
        if (id === -1) window.Asc.plugin.executeCommand("close", "");
    };

})(window, undefined);
