import { AdaptiveCards } from "@microsoft/adaptivecards-tools"
import {
  AdaptiveCardInvokeResponse,
  AdaptiveCardInvokeValue,
  CardFactory,
  MessageFactory,
  TeamsActivityHandler,
  TeamsChannelAccount,
  TeamsInfo,
  TurnContext
} from "botbuilder"
import { encode } from "html-entities"
import rawChosenCard from "./adaptiveCards/chosen.json"
import rawPickerCard from "./adaptiveCards/personPicker.json"
import rawWelcomeCard from "./adaptiveCards/welcome.json"

type Mention = {
  mentioned: TeamsChannelAccount,
  text: string,
  type: string
}

export interface DataInterface {
  chosen: string
}

export class DailyRotationBot extends TeamsActivityHandler {
  chosenObj: { chosen: string }
  availableUsersObj: { users: Promise<TeamsChannelAccount>[] }

  constructor() {
    super()

    this.chosenObj = { chosen: "John Doe" }
    this.availableUsersObj = { users: [] }

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.")

      let txt = context.activity.text
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      )
      if (removedMentionText) {
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim()
      }
      if (txt !== null && txt !== undefined) {
        switch (txt) {
          case "welcome": {
            const card =
              AdaptiveCards.declareWithoutData(rawWelcomeCard).render()
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)]
            })
            break
          }
          case "choose": {
            const card = AdaptiveCards.declare<DataInterface>(
              rawChosenCard
            ).render(this.chosenObj)
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)]
            })
            break
          }
          case "register": {
            const userCard = CardFactory.adaptiveCard(rawPickerCard)
            await context.sendActivity({ attachments: [userCard] })
            break
          }
          case "notify": {
            const msg = this.mentionActivity(context)
            await context.sendActivity(msg)
            break
          }
        }
      } else if (context.activity.value !== null && context.activity.value !== undefined) {
        const reply = await this.handleRegisterResponse(context)
        await context.sendActivity(reply)
      }

      await next()
    })
  }

  private handleRegisterResponse(context: TurnContext) {
    const combinedUserId = context.activity.value.userId
    if (typeof combinedUserId === 'string') {
      this.method(combinedUserId, context)
    }
    return this.getMembersMention()
  }

  private method(combinedUserId: string, context: TurnContext) {
    const userIds = combinedUserId.split(',')
    const members = userIds.map((userId) => TeamsInfo.getMember(context, userId))
    this.availableUsersObj.users = members
  }

  private mentionActivity(context: TurnContext) {
    const mention = this.mention(context.activity.from)
    const replyActivity = MessageFactory.text(`${ mention.text }, it's your turn today!`)
    replyActivity.entities = [mention]
    return replyActivity
  }

  private mention(member: TeamsChannelAccount) {
    return {
      mentioned: member,
      text: `<at>${encode(member.name)}</at>`,
      type: "mention"
    }
  }

  private async getMembersMention() {
    const mentions = this.availableUsersObj.users.map(async (user) => this.mention(await user))
    const mentionString = await this.getMentionsAsString(mentions)
    const replyActivity = MessageFactory.text(`Members picked : ${ mentionString }`)
    replyActivity.entities = await Promise.all(mentions)
    return replyActivity
  }

  private async getMentionsAsString(mentions: Promise<Mention>[]) {
    const mentionTexts = mentions.map(async (mention) => (await mention).text)
    return (await Promise.all(mentionTexts)).join(', ')
  }

  async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<AdaptiveCardInvokeResponse> {
    if (invokeValue.action.verb === "skip") {
      this.chosenObj.chosen = `Skipped ${ this.chosenObj.chosen }`
      const card = AdaptiveCards.declare<DataInterface>(rawChosenCard).render(
        this.chosenObj
      )
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)]
      })
      return { statusCode: 200, type: undefined, value: undefined }
    }
  }
}
