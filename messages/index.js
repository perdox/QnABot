// For more information about this template visit http://aka.ms/azurebots-node-qnamaker

"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var builder = require("botbuilder");
var botbuilder_azure = require("botbuilder-azure");
var path = require('path');
var request = require("request");

var useEmulator = (process.env.NODE_ENV == 'development');

var connector = useEmulator ? new builder.ChatConnector() : new botbuilder_azure.BotServiceConnector({
    appId: process.env['MicrosoftAppId'],
    appPassword: process.env['MicrosoftAppPassword'],
    stateEndpoint: process.env['BotStateEndpoint'],
    openIdMetadata: process.env['BotOpenIdMetadata']
});

if (useEmulator) {
    var restify = require('restify');
    var server = restify.createServer();
    server.listen(3978, function() {
        console.log('test bot endpont at http://localhost:3978/api/messages');
    });
    server.post('/api/messages', connector.listen());
} else {
    module.exports = { default: connector.listen() }
}

var request = require("request");
var entities = require("html-entities");
var qnaMakerServiceEndpoint = 'https://westus.api.cognitive.microsoft.com/qnamaker/v2.0/knowledgebases/';
var qnaApi = 'generateanswer';
var qnaTrainApi = 'train';
var htmlentities = new entities.AllHtmlEntities();
var QnAMakerRecognizer = (function () {
    function QnAMakerRecognizer(options) {
        this.options = options;
        this.kbUri = qnaMakerServiceEndpoint + this.options.knowledgeBaseId + '/' + qnaApi;
        this.kbUriForTraining = qnaMakerServiceEndpoint + this.options.knowledgeBaseId + '/' + qnaTrainApi;
        this.ocpApimSubscriptionKey = this.options.subscriptionKey;
        this.intentName = options.intentName || "qna";
        if (typeof this.options.top !== 'number') {
            this.top = 1;
        }
        else {
            this.top = this.options.top;
        }
    }
    QnAMakerRecognizer.prototype.recognize = function (context, cb) {
        var result = { score: 0.0, answers: null, intent: null };
        if (context && context.message && context.message.text) {
            console.log(context);
            console.log(context.message);
            var utterance = context.message.text;
            QnAMakerRecognizer.recognize(utterance, this.kbUri, this.ocpApimSubscriptionKey, this.top, this.intentName, function (error, result) {
                if (!error) {
                    cb(null, result);
                }
                else {
                    cb(error, null);
                }
            });
        }
    };
    QnAMakerRecognizer.recognize = function (utterance, kbUrl, ocpApimSubscriptionKey, top, intentName, callback) {
        try {
            var postBody = '{"question":"' + utterance + '", "top":' + top + '}';
            request({
                url: kbUrl,
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Ocp-Apim-Subscription-Key': ocpApimSubscriptionKey
                },
                body: postBody
            }, function (error, response, body) {
                var result;
                try {
                    if (!error) {
                        result = JSON.parse(body);
                        var answerEntities = [];
                        if (result.answers !== null && result.answers.length > 0) {
                            result.answers.forEach(function (ans) {
                                ans.score /= 100;
                                ans.answer = htmlentities.decode(ans.answer);
                                var answerEntity = {
                                    score: ans.score,
                                    entity: ans.answer,
                                    type: 'answer'
                                };
                                answerEntities.push(answerEntity);
                            });
                            result.score = result.answers[0].score;
                            result.entities = answerEntities;
                            result.intent = intentName;
                        }
                    }
                }
                catch (e) {
                    error = e;
                }
                try {
                    if (!error) {
                        callback(null, result);
                    }
                    else {
                        var m = error.toString();
                        callback(error instanceof Error ? error : new Error(m));
                    }
                }
                catch (e) {
                    console.error(e.toString());
                }
            });
        }
        catch (e) {
            callback(e instanceof Error ? e : new Error(e.toString()));
        }
    };
    return QnAMakerRecognizer;
}());

var QnAMakerTools = (function () {
    function QnAMakerTools() {
        this.lib = new builder.Library('qnaMakerTools');
        this.lib.dialog('answerSelection', [
            function (session, args) {
                var qnaMakerResult = args;
                session.dialogData.qnaMakerResult = qnaMakerResult;
                var questionOptions = [];
                qnaMakerResult.answers.forEach(function (qna) { questionOptions.push(qna.questions[0]); });
                questionOptions.push("None of the above.");
                var promptOptions = { listStyle: builder.ListStyle.button, maxRetries: 0 };
                builder.Prompts.choice(session, "Did you mean:", questionOptions, promptOptions);
            },
            function (session, results) {
                if (results && results.response && results.response.entity) {
                    var qnaMakerResult = session.dialogData.qnaMakerResult;
                    var filteredResult = qnaMakerResult.answers.filter(function (qna) { return qna.questions[0] === results.response.entity; });
                    if (filteredResult !== null && filteredResult.length > 0) {
                        var selectedQnA = filteredResult[0];
                        
                        var m = session.MakeMessage();
                        m.text = selectedQnA.answer;
                        m.speech = selectedQnA.answer;

                        session.send(m);
                        session.endDialogWithResult(selectedQnA);
                    }
                }
                else {
                    session.send("Sorry! Not able to match any of the options.");
                }
                session.endDialog();
            },
        ]);
    }
    QnAMakerTools.prototype.createLibrary = function () {
        return this.lib;
    };
    QnAMakerTools.prototype.answerSelector = function (session, options) {
        session.beginDialog('qnaMakerTools:answerSelection', options || {});
    };
    return QnAMakerTools;
}());

var QnAMakerDialog = (function (_super) {
    __extends(QnAMakerDialog, _super);
    function QnAMakerDialog(options) {
        var _this = _super.call(this) || this;
        _this.options = options;
        _this.recognizers = new builder.IntentRecognizerSet(options);
        var qnaRecognizer = _this.options.recognizers[0];
        _this.ocpApimSubscriptionKey = qnaRecognizer.ocpApimSubscriptionKey;
        _this.kbUriForTraining = qnaRecognizer.kbUriForTraining;
        _this.qnaMakerTools = _this.options.feedbackLib;
        if (typeof _this.options.qnaThreshold !== 'number') {
            _this.answerThreshold = 0.3;
        }
        else {
            _this.answerThreshold = _this.options.qnaThreshold;
        }
        if (_this.options.defaultMessage && _this.options.defaultMessage !== "") {
            _this.defaultNoMatchMessage = _this.options.defaultMessage;
        }
        else {
            _this.defaultNoMatchMessage = "No match found!";
        }
        return _this;
    }
    QnAMakerDialog.prototype.replyReceived = function (session, recognizeResult) {
        var _this = this;
        var threshold = this.answerThreshold;
        var noMatchMessage = this.defaultNoMatchMessage;
        if (!recognizeResult) {
            var locale = session.preferredLocale();
            var context = session.toRecognizeContext();
            context.dialogData = session.dialogData;
            context.activeDialog = true;
            this.recognize(context, function (error, result) {
                try {
                    if (!error) {
                        _this.invokeAnswer(session, result, threshold, noMatchMessage);
                    }
                }
                catch (e) {
                    _this.emitError(session, e);
                }
            });
        }
        else {
            this.invokeAnswer(session, recognizeResult, threshold, noMatchMessage);
        }
    };
    QnAMakerDialog.prototype.recognize = function (context, cb) {
        this.recognizers.recognize(context, cb);
    };
    QnAMakerDialog.prototype.recognizer = function (plugin) {
        this.recognizers.recognizer(plugin);
        return this;
    };
    QnAMakerDialog.prototype.invokeAnswer = function (session, recognizeResult, threshold, noMatchMessage) {
        var qnaMakerResult = recognizeResult;
        session.privateConversationData.qnaFeedbackUserQuestion = session.message.text;
        if (qnaMakerResult.score >= threshold && qnaMakerResult.answers.length > 0) {
            if (this.isConfidentAnswer(qnaMakerResult) || this.qnaMakerTools == null) {
                this.respondFromQnAMakerResult(session, qnaMakerResult);
                this.defaultWaitNextMessage(session, qnaMakerResult);
            }
            else {
                this.qnaFeedbackStep(session, qnaMakerResult);
            }
        }
        else {
            session.send(noMatchMessage);
            this.defaultWaitNextMessage(session, qnaMakerResult);
        }
    };
    QnAMakerDialog.prototype.qnaFeedbackStep = function (session, qnaMakerResult) {
        this.qnaMakerTools.answerSelector(session, qnaMakerResult);
    };
    QnAMakerDialog.prototype.respondFromQnAMakerResult = function (session, qnaMakerResult) {
        session.send(qnaMakerResult.answers[0].answer);
    };
    QnAMakerDialog.prototype.defaultWaitNextMessage = function (session, qnaMakerResult) {
        session.endDialog();
    };
    QnAMakerDialog.prototype.isConfidentAnswer = function (qnaMakerResult) {
        if (qnaMakerResult.answers.length <= 1
            || qnaMakerResult.answers[0].score >= 0.99
            || (qnaMakerResult.answers[0].score - qnaMakerResult.answers[1].score > 0.2)) {
            return true;
        }
        return false;
    };
    QnAMakerDialog.prototype.dialogResumed = function (session, result) {
        var selectedResponse = result;
        if (selectedResponse && selectedResponse.answer && selectedResponse.questions && selectedResponse.questions.length > 0) {
            var feedbackPostBody = '{"feedbackRecords": [{"userId": "' + session.message.user.id + '","userQuestion": "' + session.privateConversationData.qnaFeedbackUserQuestion
                + '","kbQuestion": "' + selectedResponse.questions[0] + '","kbAnswer": "' + selectedResponse.answer + '"}]}';
            this.recordQnAFeedback(feedbackPostBody);
        }
        this.defaultWaitNextMessage(session, { answers: [selectedResponse] });
    };
    QnAMakerDialog.prototype.recordQnAFeedback = function (body) {
        console.log(body);
        request({
            url: this.kbUriForTraining,
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                'Ocp-Apim-Subscription-Key': this.ocpApimSubscriptionKey
            },
            body: body
        }, function (error, response, body) {
            if (response.statusCode == 204) {
                console.log('Feedback sent successfully.');
            }
            else {
                console.log('error: ' + response.statusCode);
                console.log(body);
            }
        });
    };
    QnAMakerDialog.prototype.emitError = function (session, err) {
        var m = err.toString();
        err = err instanceof Error ? err : new Error(m);
        session.error(err);
    };
    return QnAMakerDialog;
}(builder.Dialog));

var bot = new builder.UniversalBot(connector);
bot.localePath(path.join(__dirname, './locale'));

var recognizer = new QnAMakerRecognizer({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
    subscriptionKey: process.env.QnASubscriptionKey});

var basicQnAMakerDialog = new QnAMakerDialog({
    recognizers: [recognizer],
                defaultMessage: 'No match! Try changing the query terms!',
                qnaThreshold: 0.3}
);


bot.dialog('/', basicQnAMakerDialog);


