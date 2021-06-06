// See Topic definition, each topic has a list of rules
// each Rule has a question and a dictionary of answers, which maps to other rules and a suggestion per answer.
// Whenever a branch reaches the end, the last rule must be END_RULE.
const peopleTopic = new Topic("Image with people", 
[
    new Rule(
        "Are the people relevant to the content?",
        {
            "no": { rule: END_RULE, suggestion: "Do not include it in the alt text." },
            "yes": {
                suggestion: "Begin by stating there are people in the image.",
                rule: new Rule(
                    "Is the number of people relevant to the context?",
                    {
                        "no": { rule: END_RULE, suggestion: "Do not include it in the alt text." },
                        "yes": {
                            suggestion: "Describe the total number of people.",
                            rule: new Rule(
                                "Is the gender/race identity relevant to the content?",
                                {
                                    "no": { rule: END_RULE, suggestion: "Do not disclose the gender/race identitiy of the people in the image." },
                                    "yes": {
                                        suggestion: "Add the general (not individual) gender/race aspects which are relevant to the context.",
                                        rule: new Rule(
                                            "Do you have permission to describe the people's identity?",
                                            {
                                                "no": { rule: END_RULE, suggestion: "Do not disclose the identities of the people." },
                                                "yes": { rule: END_RULE, suggestion: "Describe the identity of each person in a specific order." }
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

const decorativeTopic = new Topic("Decorative", [
    new Rule(
        "Is this a purely decorative image without content?",
        {
            "yes": { rule: END_RULE, suggestion: "Leave an empty alt text." },
            "no": { rule: END_RULE, suggestion: "Consider choosing another category." }
        }
    )
]);

const chartTopic = new Topic("Chart", [
    new Rule(
        "For a chart, are all of the data points important?",
        {
            "no": {rule: END_RULE, suggestion: "Shortly describe the key takeaway you want the reader to have."},
            "yes": {
                suggestion: "Consider adding a data table as well to allow table navigation.",
                rule: new Rule(
                    "Is their also a simple takeaway you want the reader to have?",
                    {
                        "no": { rule: END_RULE, suggestion: "Describe where the data can be found, only if that is not already referenced in your document." },
                        "yes": { rule: END_RULE, suggestion: "Shortly describe the key takeaway you want the reader to have. Also reference where"+ 
                        " the data table can be found, only if that is not already referenced in your document." }
                    }
                )
            }
        }
    )
]);

const TOPICS = {};
TOPICS[decorativeTopic.name] = decorativeTopic;
TOPICS[peopleTopic.name] = peopleTopic;
TOPICS[chartTopic.name] = chartTopic;