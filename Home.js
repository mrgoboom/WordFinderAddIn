
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

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

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(highlightComplexity);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    }

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
        if (xmlhttp.status == 200) {
            result = xmlhttp.responseText;
        }
        return result;
    }

    function createRegExpArray(wordText) {
        let words = wordText.split(',');
        let patterns = new Array();
        const PUNCTUATION = `!"#$%&'()*+,\-./:;<=>?@\[\\\]\^_\`{|} ~\\\\`;
        for (let i = 0; i < words.length; i++) {
            let wordPattern = new RegExp(`[${PUNCTUATION}]*${words[i]}[${PUNCTUATION}]*`);
            patterns.push(wordPattern);
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

    function highlightWords(wordPatterns) {
        Word.run( async function (context) {
            let range = context.document.getSelection();
            context.load(range, 'isEmpty');
            let searchResults = new Array();

            await context.sync()

            if (range.isEmpty) {
                range = context.document.body.getRange();
            }
            context.load(range, 'text');

            await context.sync();

            console.log(range.text);
            let words = range.text.split(/\s+/);
            let wordsMissingFromList = new Array();

            for (let i = 0; i < words.length; i++) {
                words[i] = words[i].trim();
                if (words[i] != "" && !isWordFoundOnList(words[i].toLocaleLowerCase(), wordPatterns)) {
                    console.log(`Highlight "${words[i]}"`);
                    wordsMissingFromList.push(words[i]);
                }
            }

            for (let word of wordsMissingFromList) {
                let result = range.search(word, { matchCase: true, matchWholeWord: true });
                searchResults.push(result);
                context.load(result, 'font');
            }

            await context.sync();

            for (let result of searchResults) {
                result.items[0].font.highlightColor = '#FFFF00';
            }

            await context.sync();
        }).catch(errorHandler);
    }

    const WORD_REG_EXP = /\S*\w+\S*/;

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

    function highlightSentences(maxSentenceLength) {
        Word.run(async function (context) {
            let range = context.document.getSelection();
            context.load(range, 'isEmpty');
            let searchResults = new Array();

            await context.sync();

            if (range.isEmpty) {
                range = context.document.body.getRange();
            }
            context.load(range, 'text');

            await context.sync();

            let sentences = range.text.split(/[\.\?\!]+/);
            let longSentences = new Array();

            for (let i = 0; i < sentences.length; i++) {
                if (sentences[i] == null || !WORD_REG_EXP.test(sentences[i])) {
                    continue;
                }
                console.log(sentences[i]);
                let words = sentences[i].match(WORD_REG_EXP);
                sentences[i] = sentences[i].substring(sentences[i].indexOf(words[0]));
                sentences[i] = sentences[i].trim();
                if (sentences[i] != "" && calcSentenceLength(sentences[i].toLocaleLowerCase()) > maxSentenceLength) {
                    console.log(`Highlight: "${sentences[i]}"`);
                    longSentences.push(sentences[i]);
                }
            }

            for (let sentence of longSentences) {
                let result = range.search(sentence, { matchCase: true });
                searchResults.push(result);
                context.load(result, 'font');
            }

            await context.sync();

            for (let result of searchResults) {
                result.items[0].font.highlightColor = '#00FFFF';
            }

            await context.sync();
        }).catch(errorHandler);
    }

    function highlightComplexity() {
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
            highlightSentences(maxSentenceLength);
        }

        if (doHighlightWords) {
            let wordListFile;
            if (doUseShortWordList) {
                wordListFile = "/Content/Wordlists/shortlist.txt";
            } else {
                wordListFile = "/Content/Wordlists/longlist.txt";
            }
            let wordListText = loadFile(wordListFile);
            let wordListPatterns = createRegExpArray(wordListText);
            highlightWords(wordListPatterns);
        }
    }

    function hightlightLongestWord() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
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
