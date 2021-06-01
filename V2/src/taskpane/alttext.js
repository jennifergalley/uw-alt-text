class Rule {
    /* 
    question: A string with the question to be prompted to the user.
    answers:
    {
        <answer string>: {
            suggestion: suggestion based on <answer string>
            rule: Rule object which should get triggered after selecting <answer string>
        }
    }
    */
    constructor(question, answers) {
        this.question = question;
        this.answers = answers;
    }

    nextRule(answer) {
        if (this.answers[answer]) {
            return this.answers[answer];
        }
        
        throw "Invalid answer";
    }
}

END_RULE = new Rule("END_RULE", [], []); // Use this as a base case

class Topic {
    constructor(name, rules) {
        this.name = name;
        this.rules = rules;
    }
  

    // https://stackoverflow.com/a/44878621/1391660
    toJSON() {
        return Object.getOwnPropertyNames(this).reduce((a, b) => {
            a[b] = this[b];
            return a;
        }, {});
    }
}

class Engine {
    constructor(topic) {
        this.topic = topic;
        this.ruleIdx = 0;
        this.currentRule = this.topic.rules[this.ruleIdx];
    }

    nextQuestion() {
        return {
            question: this.currentRule.question,
            answers: Object.keys(this.currentRule.answers)
        }
    }

    receiveAnswer(answer) {
        let aux = this.currentRule.nextRule(answer);
        this.currentRule = aux.rule;
        if (this.currentRule == END_RULE) {
            this.ruleIdx++;
            if (!this.done()) {
                this.currentRule = this.topic.rules[this.ruleIdx];
            }
        }

        return {
            suggestion: aux.suggestion,
            question: this.nextQuestion()
        };
    }

    done() {
        return this.ruleIdx >= this.topic.rules.length;
    }

}

let engine = null;
let suggestions = [];

function onTopicChosen(event) {
    let topic = TOPICS[event.target.value];
    if (topic) {
        document.getElementById("current-question-next-button").disabled = false;
        engine = new Engine(topic);
        suggestions = [];
        let q = engine.nextQuestion();
        updateQuestionUI(q);
    } else {
        document.getElementById("current-question-div").style.display = "none";
        document.getElementById("current-question-next-button").disabled = true;
    }
}

function nextQuestion() {
    let dropdown = document.getElementById("current-question-dropdown");
    let answer = dropdown.options[dropdown.selectedIndex].value;
    let question = engine.nextQuestion().question;
    let ans = engine.receiveAnswer(answer);
    suggestions.push([question, answer, ans.suggestion]);
    if (!engine.done()) {
        updateQuestionUI(ans.question);
    } else {
        console.log("DONE ", suggestions);
        document.getElementById("current-question-div").style.display = "none";
        document.getElementById("suggestions-div").style.display = "flex";
        let suggestionList = document.getElementById("suggestions-list");
        while (suggestionList.firstChild) {
            suggestionList.firstChild.remove()
        }
        for (const [question, answer, suggestion] of suggestions) {
            let htmlStr = `
            <li class="ms-ListItem is-selectable" tabindex="0">
                <span class="ms-ListItem-primaryText">${question}</span> 
                <span class="ms-ListItem-secondaryText">${answer}</span> 
                <span class="ms-ListItem-tertiaryText">${suggestion}</span> 
                <div class="ms-ListItem-selectionTarget"></div>
            </li>
            `;

            const fragment = document.createRange().createContextualFragment(htmlStr);

            suggestionList.appendChild(fragment);
            console.log(question, answer, suggestion);
        }
        suggestions = [];
    }
}

function updateQuestionUI(q) {
    document.getElementById("current-question-div").style.display = "flex";
    document.getElementById("suggestions-div").style.display = "none";
    document.getElementById("current-question-text").textContent = q.question;
    let dropdown = document.getElementById("current-question-dropdown");
    while (dropdown.firstChild) {
        dropdown.firstChild.remove()
    }
    for (const answer of q.answers) {
        console.log("ADDING " + answer)
        var option = document.createElement("option");
        option.appendChild(document.createTextNode(answer));
        dropdown.appendChild(option);
    }

    // ======= Office ui fabric core insanity. =======
    // See https://itgeneralisten.wordpress.com/2017/02/14/the-office-fabric-ui-dropdown/
    // and https://social.msdn.microsoft.com/Forums/en-US/476da803-e4f6-46af-ad38-2fb22223991e/how-do-i-reinitialize-a-office-ui-fabric-dropdown?forum=sharepointdevelopment
    let dropdownDiv = document.getElementById("current-question-dropdown-div");
    let child = dropdownDiv.querySelector(".ms-Dropdown-title");
    if (child != null) {
        dropdownDiv.removeChild(child);
    }
    
    child = dropdownDiv.querySelector(".ms-Dropdown.item");
    if (child != null) {
        dropdownDiv.removeChild(child);
    }
    
    new fabric['Dropdown'](dropdownDiv);
}

window.addEventListener("load", function(){

    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();
    document.getElementById("current-question-next-button").disabled = true;
    let dropdown = document.getElementById("topics-dropdown");
    dropdown.onchange = onTopicChosen;
    $("#update-altext-container").hide();
    for (const topic of Object.values(TOPICS)) {
        console.log("Adding " + topic.name);
        var option = document.createElement("option");
        option.appendChild(document.createTextNode(topic.name));
        dropdown.appendChild(option);
    }

    // Required by fabric ui core dropdown docs (after dynamically adding data): https://developer.microsoft.com/en-us/fabric-js/components/dropdown/dropdown
    var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
    for (var i = 0; i < DropdownHTMLElements.length; ++i) {
        var Dropdown = new fabric['Dropdown'](DropdownHTMLElements[i]);
    }
    
    let nextButton = document.getElementById("current-question-next-button");
    new fabric['Button'](nextButton, nextQuestion);

    var ListElements = document.querySelectorAll(".ms-List");
    for (var i = 0; i < ListElements.length; i++) {
        new fabric['List'](ListElements[i]);
    }
});

