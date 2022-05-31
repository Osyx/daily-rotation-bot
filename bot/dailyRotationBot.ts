import { AdaptiveCards } from "@microsoft/adaptivecards-tools"
import {
  Activity,
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

type ChosenOutput = {
  cardActivity: Partial<Activity>,
  pingActivity: Partial<Activity>
}

type SaveData = {
  chosenIndex: number,
  registeredUserIds: string[]
}

export interface ChosenDataInterface {
  chosen: string
}

const DEFAULTS: SaveData = { registeredUserIds: [], chosenIndex: undefined }
const CONVERSATION_DATA_PROPERTY = 'conversationData'

export class DailyRotationBot extends TeamsActivityHandler {

  chosenObj: ChosenDataInterface
  conversationState: ConversationState
  conversationDataAccessor: StatePropertyAccessor

  constructor(conversationState: ConversationState) {
    super()

    this.conversationState = conversationState
    this.conversationDataAccessor = conversationState.createProperty(CONVERSATION_DATA_PROPERTY)

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.")

      const conversationData: SaveData = await this.conversationDataAccessor.get(context, DEFAULTS)

      let txt: string = context.activity.text
      const removedMentionText: string = TurnContext.removeRecipientMention(
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
            const chosenOutput: ChosenOutput = await this.handleChoose(context, conversationData)
            await context.sendActivity(chosenOutput.cardActivity)
            if (chosenOutput.pingActivity !== undefined) {
              await context.sendActivity(chosenOutput.pingActivity)
            }
            break
          }
          case "register": {
            const userCard = CardFactory.adaptiveCard(rawPickerCard)
            await context.sendActivity({ attachments: [userCard] })
            break
          }
          case "users": {
            if (conversationData.registeredUserIds) {
              const usersActivity = await this.handleRegisteredUsers(context, conversationData)
              await context.sendActivity(usersActivity)
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

  private async handleChoose(context: TurnContext, conversationData: SaveData): Promise<ChosenOutput> {
    const chosenIndex: number = conversationData.chosenIndex
    const registeredUserIds: string[] = conversationData.registeredUserIds
    let member: TeamsChannelAccount
    if (chosenIndex !== undefined) {
      let nextChosenIndex = chosenIndex + 1
      nextChosenIndex = registeredUserIds.length > nextChosenIndex
        ? nextChosenIndex
        : 0
      conversationData.chosenIndex = nextChosenIndex
      member = await TeamsInfo.getMember(context, registeredUserIds[nextChosenIndex])
      this.chosenObj = { chosen: member.name }
    } else {
      this.chosenObj = { chosen: "John Doe" }
    }

    return {
      cardActivity: this.constructChosenActivity(member),
      pingActivity: this.constructPingActivity(member)
    }
  }

  private constructChosenActivity(member: TeamsChannelAccount): Partial<Activity> {
    const chosenActivity: Partial<Activity>  = {}
    const renderedChosenCard = AdaptiveCards.declare<ChosenDataInterface>(rawChosenCard).render(this.chosenObj)
    chosenActivity.attachments = [CardFactory.adaptiveCard(renderedChosenCard)]
    return chosenActivity
  }

  private constructPingActivity(member: TeamsChannelAccount): Partial<Activity> {
    if (member === undefined) {
      return undefined
    }
    const mention = this.mention(member)
    const chosenActivity: Partial<Activity>  = MessageFactory.text(`Your turn, ${ mention.text }!`)
    chosenActivity.entities = [mention]
    return chosenActivity
  }

  private async handleRegisteredUsers(context: TurnContext, conversationData: SaveData): Promise<Partial<Activity>> {
    const userIds = conversationData.registeredUserIds
    const members = userIds.map((userId) => TeamsInfo.getMember(context, userId))
    return this.getSavedMembersActivity(members)
  }

  private async getSavedMembersActivity(members: Promise<TeamsChannelAccount>[]): Promise<Partial<Activity>> {
    const promisedMentions = members.map(async (user) => this.mention(await user))
    const mentions = await Promise.all(promisedMentions)
    const mentionTexts = mentions.map((mention) => mention.text)
    const mentionText = mentionTexts.join(', ')
    const savedMembersActivity = MessageFactory.text(`Current rotation members: ${ mentionText.length === 0 ? 'None!' : mentionText }`)
    savedMembersActivity.entities = mentions
    return savedMembersActivity
  }

  private mention(member: TeamsChannelAccount) {
    return {
      mentioned: member,
      text: `<at>${ encode(member.name) }</at>`,
      type: "mention"
    }
  }

  private handleRegisterResponse(context: TurnContext, conversationData: SaveData): void {
    const combinedUserId = context.activity.value.userId
    if (typeof combinedUserId === 'string') {
      const userIds = combinedUserId.split(',')
      conversationData.chosenIndex = 0
      conversationData.registeredUserIds = userIds
    }
  }

  async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<AdaptiveCardInvokeResponse> {
    if (invokeValue.action.verb === "skip") {
      const conversationData: SaveData = await this.conversationDataAccessor.get(context, DEFAULTS)
      const chosenOutput: ChosenOutput = await this.handleChoose(context, conversationData)
      const updatedActivity = chosenOutput.cardActivity
      updatedActivity.id = context.activity.replyToId
      updatedActivity.type = 'message'
      await context.updateActivity(updatedActivity)
      if (chosenOutput.pingActivity !== undefined) {
        await context.sendActivity(chosenOutput.pingActivity)
      }
      return { statusCode: 200, type: undefined, value: undefined }
    }
  }

  async run(context: TurnContext): Promise<void> {
    await super.run(context)
    await this.conversationState.saveChanges(context, false)
  }
}
