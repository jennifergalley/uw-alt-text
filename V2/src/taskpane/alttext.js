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

let peopleTopic = new Topic("Image with people", 
    [
        new Rule(
            "Is the people relevant to the content at hand?",
            {
                "no": { rule: END_RULE, suggestion: "Do not include this information in the alt text." },
                "yes": {
                    suggestion: "Begin stating there is people in the image.",
                    rule: new Rule(
                        "Is the number of people relevant to the context?",
                        {
                            "no": { rule: END_RULE, suggestion: "Do not include this information in the alt text." },
                            "yes": {
                                suggestion: "Describe the total number of people in the image, with emphasis on the people who are relevant to the context.",
                                rule: new Rule(
                                    "Is the gender/race identity relevant to the message you are trying to convey?",
                                    {
                                        "no": { rule: END_RULE, suggestion: "Do not disclose the identitiy of the people in the image." },
                                        "yes": {
                                            suggestion: "Add the general (not indidividual) gender/race aspects which are relevant to the context",
                                            rule: new Rule(
                                                "Do you have an explicit acknowledgement of the people in the picture to describe their identity?",
                                                {
                                                    "no": { rule: END_RULE, suggestion: "Do not disclose the identitiy of the people in the image." },
                                                    "yes": { rule: END_RULE, suggestion: "Describe accurately the identity of each person from a specific order, say left to right." }
                                                }
                                            )
                                        }
                                    }
                                )
                            }
                        }
                    )
                }
            }
        )
    ]
);

let decorativeTopic = new Topic("Decorative", [
    new Rule(
        "Is this a purely decorative image without content?",
        {
            "yes": { rule: END_RULE, suggestion: "Leave an empty alt text." },
            "no": { rule: END_RULE, suggestion: "Consider choosing another category" }
        }
    )
]);

let TOPICS = {};
TOPICS[decorativeTopic.name] = decorativeTopic;
TOPICS[peopleTopic.name] = peopleTopic;
let engine = null;
let suggestions = [];

function onTopicChosen(event) {
    let topic = TOPICS[event.target.value];
    engine = new Engine(topic);
    suggestions = [];
    let q = engine.nextQuestion();
    updateQuestionUI(q);
}

function nextQuestion() {
    let dropdown = document.getElementById("current-question-dropdown");
    let answer = dropdown.options[dropdown.selectedIndex].value;
    let ans = engine.receiveAnswer(answer);
    suggestions.push(ans.suggestion);
    if (!engine.done()) {
        updateQuestionUI(ans.question);
    } else {
        console.log("DONE ", suggestions);
        document.getElementById("current-question-div").style.display = "none";
        document.getElementById("suggestions-div").style.display = "flex";
        let paragraph = document.getElementById("suggestions-paragraph");
        for (let suggestion of suggestions) {
            paragraph.textContent += suggestion + "\n";
        }
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
}

window.addEventListener("load", function(){
    document.getElementById("current-question-next-button").onclick = nextQuestion;
    let dropdown = document.getElementById("topics-dropdown");
    dropdown.onchange = onTopicChosen;
    for (const topic of Object.values(TOPICS)) {
        console.log("Adding " + topic.name);
        var option = document.createElement("option");
        option.appendChild(document.createTextNode(topic.name));
        dropdown.appendChild(option);
    }
});

