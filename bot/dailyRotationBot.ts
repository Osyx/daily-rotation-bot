import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import {
  AdaptiveCardInvokeResponse, AdaptiveCardInvokeValue, CardFactory, MessageFactory, TeamsActivityHandler, TurnContext
} from "botbuilder";
import { encode } from "html-entities";
import rawChosenCard from "./adaptiveCards/chosen.json";
import rawPickerCard from "./adaptiveCards/personPicker.json";
import rawWelcomeCard from "./adaptiveCards/welcome.json";

export interface DataInterface {
  chosen: string;
}

export class DailyRotationBot extends TeamsActivityHandler {
  chosenObj: { chosen: string };
  availableUsersObj: { users: string[] };

  constructor() {
    super();

    this.chosenObj = { chosen: "John Doe" };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }
      if (txt != null) {
        switch (txt) {
          case "welcome": {
            const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
            await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
            break;
          }
          case "choose": {
            const card = AdaptiveCards.declare<DataInterface>(rawChosenCard).render(this.chosenObj);
            await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
            break;
          }
          case "register": {
            const userCard = CardFactory.adaptiveCard(rawPickerCard);
            await context.sendActivity({ attachments: [userCard] });
            break;
          }
          case 'notify': {
            const msg = this.mentionActivityAsync(context);
            await context.sendActivity(msg);
            break;
          }
        }
    } else if (context.activity.value != null) {
      for (const userId in context.activity.value) {
        this.availableUsersObj.users.push(userId)
      }
      await context.sendActivity(`Members picked : ${ this.availableUsersObj.users }` );
    }

      await next();
    });
  }

  mentionActivityAsync(context: TurnContext) {
   const mention = {
        mentioned: context.activity.from,
        text: `<at>${ encode(context.activity.from.name) }</at>`,
        type: 'mention'
    };

    const replyActivity = MessageFactory.text(`${ mention.text }, it's your turn today!`);
    replyActivity.entities = [mention];
    
    return replyActivity;
  }

  async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<AdaptiveCardInvokeResponse> {
    if (invokeValue.action.verb === "skip") {
      this.chosenObj.chosen = `Skipped ${ this.chosenObj.chosen }`;
      const card = AdaptiveCards.declare<DataInterface>(rawChosenCard).render(this.chosenObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }
}

