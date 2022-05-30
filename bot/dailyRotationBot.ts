import { AdaptiveCards } from "@microsoft/adaptivecards-tools"
import {
  AdaptiveCardInvokeResponse,
  AdaptiveCardInvokeValue,
  CardFactory,
  ConversationState,
  MessageFactory,
  StatePropertyAccessor,
  TeamsActivityHandler,
  TeamsChannelAccount,
  TeamsInfo,
  TurnContext
} from "botbuilder"
import { encode } from "html-entities"
import rawChosenCard from "./adaptiveCards/chosen.json"
import rawPickerCard from "./adaptiveCards/personPicker.json"
import rawPingCard from "./adaptiveCards/ping.json"


type SaveData = {
  registeredUserIds: string[]
}

export interface DataInterface {
  chosen: string
}

const CONVERSATION_DATA_PROPERTY = 'conversationData'

export class DailyRotationBot extends TeamsActivityHandler {

  chosenObj: { chosen: string }
  conversationState: ConversationState
  conversationDataAccessor: StatePropertyAccessor

  constructor(conversationState: ConversationState) {
    super()

    this.chosenObj = { chosen: "John Doe" }
    this.conversationState = conversationState
    this.conversationDataAccessor = conversationState.createProperty(CONVERSATION_DATA_PROPERTY)

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.")

      const conversationData = await this.conversationDataAccessor.get(context, { registeredUserIds: [] })

      let txt = context.activity.text
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
        )
        if (removedMentionText) {
          txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim()
        }
        if (txt !== null && txt !== undefined) {
          switch (txt) {
          case "ping": {
            const card = AdaptiveCards.declareWithoutData(rawPingCard).render()
            await context.sendActivity({
              attachments: [CardFactory.adaptiveCard(card)]
            })
            break
          }
          case "choose": {
            const card = AdaptiveCards.declare<DataInterface>(rawChosenCard).render(this.chosenObj)
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
          case "users": {
            if (conversationData.registeredUserIds) {
              const msg = await this.handleRegisteredUsers(context, conversationData)
              await context.sendActivity(msg)
            }
            break
          }
        }
      } else if (context.activity.value !== null && context.activity.value !== undefined) {
        this.handleRegisterResponse(context, conversationData)
      }

      await next()
    })
  }

  private async handleRegisteredUsers(context: TurnContext, conversationData: SaveData) {
    const userIds = conversationData.registeredUserIds
    const members = userIds.map((userId) => TeamsInfo.getMember(context, userId))
    return this.getSavedMembersMention(members)
  }

  private async getSavedMembersMention(members: Promise<TeamsChannelAccount>[]) {
    const promisedMentions = members.map(async (user) => this.mention(await user))
    const mentions = await Promise.all(promisedMentions)
    const mentionTexts = mentions.map((mention) => mention.text)
    const mentionText = mentionTexts.join(', ')
    const replyActivity = MessageFactory.text(`Saved members: ${ mentionText }`)
    replyActivity.entities = mentions
    return replyActivity
  }

  private mention(member: TeamsChannelAccount) {
    return {
      mentioned: member,
      text: `<at>${ encode(member.name) }</at>`,
      type: "mention"
    }
  }

  private handleRegisterResponse(context: TurnContext, conversationData: SaveData) {
    const combinedUserId = context.activity.value.userId
    if (typeof combinedUserId === 'string') {
      const userIds = combinedUserId.split(',')
      conversationData.registeredUserIds = userIds
    }
  }

  async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<AdaptiveCardInvokeResponse> {
    if (invokeValue.action.verb === "skip") {
      this.chosenObj.chosen = `Skipped ${ this.chosenObj.chosen }`
      const card = AdaptiveCards.declare<DataInterface>(rawChosenCard).render(this.chosenObj)
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)]
      })
      return { statusCode: 200, type: undefined, value: undefined }
    }
  }

  async run(context: TurnContext) {
    await super.run(context)
    await this.conversationState.saveChanges(context, false)
}
}
