const { TeamsActivityHandler, TurnContext } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      // Respond to basic greetings
      if (txt.includes("hello") || txt.includes("hi")) {
        await context.sendActivity("Hello I am Faizan! How can I assist you today?");
      } else if (txt.includes("help")) {
        await context.sendActivity("Sure, I'm here to help! What do you need assistance with?");
      } else {
        // Default response
        await context.sendActivity("Hello World!");
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity("Hello world! This is me, Muhammad Faizan. How can I help you?");
          break;
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;
