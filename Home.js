
(function () {
    "use strict";

    var messageBanner;
    let numSentences;
    let numLongSentences;
    let numWords;
    let numMissingWords;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            numWords = 0;
            numMissingWords = 0;
            numSentences = 0;
            numLongSentences = 0;

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");

                $('#highlight-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("WordFinder highlights selected words not on the specified list or sentences longer than the specified length.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights complex words and long sentences.");

            //loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(highlightComplexity);

            $('#clear-button-text').text("Clear");
            $('#clear-button-desc').text("Clears existing highlighting.");

            $('#clear-button').click(clearHighlighting);
        });
    };

    function getFormInfo() {
        let form = document.forms["settings"];
        return form;
    }

    function loadFile(inputFile) {
        if (!inputFile) {
            return;
        }
        console.log("Filename: " + inputFile);
        let result;
        let xmlhttp = new XMLHttpRequest();
        xmlhttp.open("GET", inputFile, false);
        xmlhttp.send();

        //let debugCell = document.getElementById("debug-cell");
        //debugCell.innerHTML = "" + xmlhttp.status;

        if (xmlhttp.status == 200) {
            result = xmlhttp.responseText;
        }
        return result;
    }

    function createRegExpArray(wordText) {
        let words = wordText.split(',');
        let patterns = new Array();
        const PUNCTUATION = `????????????!"#$%&'()*+,\\-.\\/:;<=>?@[\\]^_\`{|}~\\\\`;
        for (let i = 0; i < words.length; i++) {
            words[i] = words[i].toLowerCase().trim();
            words[i] = words[i].replaceAll("'", "['??????]");
            if (words[i] != "") {
                let wordPattern = new RegExp(`[${PUNCTUATION}]*${words[i]}[${PUNCTUATION}]*`);
                patterns.push(wordPattern);
            }
        }
        return patterns;
    }

    function isWordFoundOnList(word, patterns) {
        for (let pattern of patterns) {
            let matchArray = word.match(pattern);
            if (matchArray != null && word.length == matchArray[0].length) {
                return true;
            }
        }
        return false;
    }

    function highlightWords(wordPatterns, callback) {
        Word.run(async function (context) {
            let range = context.document.getSelection();
            context.load(range, 'isEmpty');
            let searchResults = new Array();
            numMissingWords = 0;
            numWords = 0;

            await context.sync()

            if (range.isEmpty) {
                range = context.document.body.getRange();
            }
            context.load(range, 'text');

            await context.sync();

            //console.log(range.text);
            let words = range.text.split(/\s+/);
            let wordsMissingFromList = new Array();

            for (let i = 0; i < words.length; i++) {
                words[i] = words[i].trim();
                if (words[i] != "") {
                    numWords++;
                    if (!isWordFoundOnList(words[i].toLowerCase(), wordPatterns)) {
                        //console.log(`Highlight "${words[i]}"`);
                        numMissingWords++;
                        wordsMissingFromList.push(words[i]);
                    }
                }
            }

            for (let word of wordsMissingFromList) {
                let result = range.search(word, { matchCase: true, matchWholeWord: true });
                searchResults.push(result);
                context.load(result, 'font');
            }

            await context.sync();

            for (let result of searchResults) {
                for (let i = 0; i < result.items.length; i++) {
                    result.items[i].font.highlightColor = '#FFFF00';
                }
            }
            console.log("Highlighted words.");

            await context.sync();

            callback();
        }).catch(errorHandler);
    }

    const WORD_REG_EXP = /\S*[\w'\-]+\S*/;

    function calcSentenceLength(sentence) {
        let len = 0;
        let words = sentence.split(/\s+/);

        for (let word of words) {
            if (WORD_REG_EXP.test(word)) {
                len++;
            }
        }
        return len;
    }

    function highlightSentences(maxSentenceLength, callback) {
        Word.run(async function (context) {
            let range = context.document.getSelection();
            context.load(range, 'isEmpty');
            let searchResults = new Array();
            numSentences = 0;
            numLongSentences = 0;

            await context.sync();

            if (range.isEmpty) {
                range = context.document.body.getRange();
            }

            context.load(range, 'paragraphs');

            await context.sync();

            let sentences = new Array();

            for (let i = 0; i < range.paragraphs.items.length; i++) {
                if (range.paragraphs.items[i].text) {
                    let newSentences = range.paragraphs.items[i].text.split(/[\.\?\!]+/);
                    sentences = sentences.concat(newSentences);
                }
            }

            let longSentences = new Array();

            for (let i = 0; i < sentences.length; i++) {
                if (sentences[i] == null || !WORD_REG_EXP.test(sentences[i])) {
                    continue;
                }
                //console.log(sentences[i]);
                let words = sentences[i].match(WORD_REG_EXP);
                sentences[i] = sentences[i].substring(sentences[i].indexOf(words[0]));
                sentences[i] = sentences[i].trim();
                if (sentences[i] != "") {
                    numSentences++;
                    if (calcSentenceLength(sentences[i].toLowerCase()) > maxSentenceLength) {
                        //console.log(`Highlight: "${sentences[i]}"`);
                        numLongSentences++;
                        longSentences.push(sentences[i]);
                    }
                }
            }

            for (let sentence of longSentences) {
                if (sentence.length < 256) {
                    let result = range.search(sentence, { matchCase: true });
                    searchResults.push(result);
                    context.load(result, 'font');
                } else {
                    let startFragment;
                    let endFragment = sentence;
                    while (endFragment.length > 510) {
                        startFragment = endFragment.substring(0, 255);
                        endFragment = endFragment.substring(255);

                        let result = range.search(startFragment, { matchCase: true });
                        searchResults.push(result);
                        context.load(result, 'font');
                    }
                    let halfIndex = Math.ceil(endFragment.length / 2);
                    startFragment = endFragment.substring(0, halfIndex);
                    endFragment = endFragment.substring(halfIndex);

                    let startResult = range.search(startFragment, { matchCase: true });
                    searchResults.push(startResult);
                    context.load(startResult, 'font');

                    let endResult = range.search(endFragment, { matchCase: true });
                    searchResults.push(endResult);
                    context.load(endResult, 'font');
                }
            }

            await context.sync();

            for (let result of searchResults) {
                for (let i = 0; i < result.items.length; i++) {
                    result.items[i].font.highlightColor = '#00FFFF';
                }
            }
            console.log("Highlighted sentences.");

            await context.sync();

            callback();
        }).catch(errorHandler);
    }

    function highlightComplexity() {
        clearHighlighting();
        let settings = getFormInfo();

        let doUseShortWordList = settings["short"].checked;
        console.log(`Use short word list?: ${doUseShortWordList}`);

        let doHighlightWords = settings["highlightWords"].checked;
        console.log(`Highlight words?: ${doHighlightWords}`);

        let maxSentenceLength = +settings["maxSentenceLength"].value > 0 ? +settings["maxSentenceLength"].value : 8;
        console.log(`Maximum sentence length: ${maxSentenceLength}`);

        let doHighlightSentences = settings["highlightSentences"].checked;
        console.log(`Highlight sentences?: ${doHighlightSentences}`);

        if (doHighlightSentences) {
            highlightSentences(maxSentenceLength, function () {
                let numSentencesCell = document.getElementById("num-sentences-cell");
                numSentencesCell.innerHTML = "" + numSentences;

                let percentSentencesCell = document.getElementById("percent-sentences-cell");
                percentSentencesCell.innerHTML = `${Math.round(100 * numLongSentences / numSentences)}%`;
            });
        }

        if (doHighlightWords) {
            //let debugCell = document.getElementById("debug-cell");
            //debugCell.innerHTML = document.location.pathname; //  /WordFinderAddIn/Home.html
            let wordListFile;
            if (doUseShortWordList) {
                wordListFile = "/WordFinderAddIn/Content/Wordlists/shortlist.txt";
            } else {
                wordListFile = "/WordFinderAddIn/Content/Wordlists/longlist.txt";
            }
            let wordListText = loadFile(wordListFile);
            let wordListPatterns = createRegExpArray(wordListText);
            highlightWords(wordListPatterns, function () {
                let numWordsCell = document.getElementById("num-words-cell");
                numWordsCell.innerHTML = "" + numWords;

                let percentWordsCell = document.getElementById("percent-words-cell");
                percentWordsCell.innerHTML = `${Math.round(100 * (numWords - numMissingWords) / numWords)}%`;
            });
        }

        if (!doHighlightWords) {
            let numWordsCell = document.getElementById("num-words-cell");
            numWordsCell.innerHTML = "N/A";

            let percentWordsCell = document.getElementById("percent-words-cell");
            percentWordsCell.innerHTML = "N/A";
        }
        if (!doHighlightSentences) {
            let numSentencesCell = document.getElementById("num-sentences-cell");
            numSentencesCell.innerHTML = "N/A";

            let percentSentencesCell = document.getElementById("percent-sentences-cell");
            percentSentencesCell.innerHTML = "N/A";
        }
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    function clearHighlighting() {
        Word.run(function (context) {
            let range = context.document.getSelection();
            context.load(range, 'isEmpty');

            return context.sync()
                .then(function () {
                    if (range.isEmpty) {
                        range = context.document.body.getRange();
                    }
                    context.load(range, 'font');
                })
                .then(context.sync)
                .then(function () {
                    range.font.highlightColor = null;
                    console.log("Highlight color cleared.");
                })
                .then(context.sync)
        }).catch(errorHandler);
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
