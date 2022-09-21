import { AdaptiveCards } from "@microsoft/adaptivecards-tools"
import {
  Activity,
  AdaptiveCardInvokeResponse,
  AdaptiveCardInvokeValue,
  CardFactory,
  ConversationState,
  MessageFactory,
  StatePropertyAccessor,
  TaskModuleRequest,
  TeamsActivityHandler,
  TeamsChannelAccount,
  TeamsInfo,
  TurnContext
} from "botbuilder"
import { encode } from "html-entities"
import rawChosenCard from "./adaptiveCards/chosen.json"
import rawPickerCard from "./adaptiveCards/personPicker.json"
import rawPingCard from "./adaptiveCards/ping.json"
import rawTaskModuleCard from "./adaptiveCards/taskModule.json"
import { TaskManager } from "./taskManager"

type ChosenOutput = {
  cardActivity: Partial<Activity>
  pingActivity: Partial<Activity>
}

type SaveData = {
  chosenName: string
  chosenIndex: number
  registeredUserIds: string[]
}

export interface ChosenDataInterface {
  chosen: string
}

const DEFAULTS: SaveData = {
  registeredUserIds: [],
  chosenIndex: undefined,
  chosenName:
    'There\'s no-one to choose! Register members with the "register" command.'
}
const CONVERSATION_DATA_PROPERTY = "conversationData"

export class DailyRotationBot extends TeamsActivityHandler {
  chosenObj: ChosenDataInterface
  conversationState: ConversationState
  conversationDataAccessor: StatePropertyAccessor
  taskManager: TaskManager

  constructor(conversationState: ConversationState) {
    super()

    this.conversationState = conversationState
    this.conversationDataAccessor = conversationState.createProperty(
      CONVERSATION_DATA_PROPERTY
    )

    this.taskManager = new TaskManager()

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.")

      const conversationData: SaveData =
        await this.conversationDataAccessor.get(context, DEFAULTS)

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
            const chosenOutput: ChosenOutput =
              await DailyRotationBot.handleChoose(
                context,
                conversationData,
                this.chosenObj
              )
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
              const usersActivity = await this.handleRegisteredUsers(
                context,
                conversationData
              )
              await context.sendActivity(usersActivity)
            }
            break
          }
          case "schedule": {
            const userCard = CardFactory.adaptiveCard(rawTaskModuleCard)
            await context.sendActivity({ attachments: [userCard] })
            break
          }
        }
      } else if (
        context.activity.value !== null &&
        context.activity.value !== undefined
      ) {
        this.handleRegisterResponse(context, conversationData)
      }

      await next()
    })
  }

  static async handleChoose(
    context: TurnContext,
    conversationData: SaveData,
    chosenObj: ChosenDataInterface
  ): Promise<ChosenOutput> {
    const chosenIndex: number = conversationData.chosenIndex
    const registeredUserIds: string[] = conversationData.registeredUserIds
    let member: TeamsChannelAccount
    if (chosenIndex !== undefined) {
      conversationData.chosenIndex =
        registeredUserIds.length > chosenIndex + 1 ? chosenIndex + 1 : 0
      member = await TeamsInfo.getMember(
        context,
        registeredUserIds[conversationData.chosenIndex]
      )
      chosenObj = { chosen: member.name }
    } else {
      chosenObj = {
        chosen:
          'There\'s no-one to choose! Register members with the "register" command.'
      }
    }

    return {
      cardActivity: this.constructChosenActivity(member, chosenObj),
      pingActivity: this.constructPingActivity(member)
    }
  }

  private static constructChosenActivity(
    member: TeamsChannelAccount,
    chosenObj: ChosenDataInterface
  ): Partial<Activity> {
    const chosenActivity: Partial<Activity> = {}
    const renderedChosenCard =
      AdaptiveCards.declare<ChosenDataInterface>(rawChosenCard).render(
        chosenObj
      )
    chosenActivity.attachments = [CardFactory.adaptiveCard(renderedChosenCard)]
    return chosenActivity
  }

  private static constructPingActivity(
    member: TeamsChannelAccount
  ): Partial<Activity> {
    if (member === undefined) {
      return undefined
    }
    const mention = this.mention(member)
    const chosenActivity: Partial<Activity> = MessageFactory.text(
      `Your turn, ${mention.text}!`
    )
    chosenActivity.entities = [mention]
    return chosenActivity
  }

  private async handleRegisteredUsers(
    context: TurnContext,
    conversationData: SaveData
  ): Promise<Partial<Activity>> {
    const userIds = conversationData.registeredUserIds
    const members = userIds.map((userId) =>
      TeamsInfo.getMember(context, userId)
    )
    return this.getSavedMembersActivity(members)
  }

  private async getSavedMembersActivity(
    members: Promise<TeamsChannelAccount>[]
  ): Promise<Partial<Activity>> {
    const promisedMentions = members.map(async (user) =>
      DailyRotationBot.mention(await user)
    )
    const mentions = await Promise.all(promisedMentions)
    const mentionTexts = mentions.map((mention) => mention.text)
    const mentionText = mentionTexts.join(", ")
    const savedMembersActivity = MessageFactory.text(
      `Current rotation members: ${
        mentionText.length === 0 ? "None!" : mentionText
      }`
    )
    savedMembersActivity.entities = mentions
    return savedMembersActivity
  }

  private static mention(member: TeamsChannelAccount) {
    return {
      mentioned: member,
      text: `<at>${encode(member.name)}</at>`,
      type: "mention"
    }
  }

  private handleRegisterResponse(
    context: TurnContext,
    conversationData: SaveData
  ): void {
    const combinedUserId = context.activity.value.userId
    if (typeof combinedUserId === "string") {
      const userIds = combinedUserId.split(",")
      conversationData.chosenIndex = -1
      conversationData.registeredUserIds = userIds
    }
  }

  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    if (invokeValue.action.verb === "skip") {
      const conversationData: SaveData =
        await this.conversationDataAccessor.get(context, DEFAULTS)
      const chosenOutput: ChosenOutput = await DailyRotationBot.handleChoose(
        context,
        conversationData,
        this.chosenObj
      )
      const updatedActivity = chosenOutput.cardActivity
      updatedActivity.id = context.activity.replyToId
      updatedActivity.type = "message"
      await context.updateActivity(updatedActivity)
      if (chosenOutput.pingActivity !== undefined) {
        await context.sendActivity(chosenOutput.pingActivity)
      }
      return { statusCode: 200, type: undefined, value: undefined }
    }
  }

  // Handle task module fetch.
  handleTeamsTaskModuleFetch(
    context: TurnContext,
    taskModuleRequest: TaskModuleRequest
  ) {
    return this.taskManager.handleFetch(taskModuleRequest)
  }

  // Handle task module submit action.
  async handleTeamsTaskModuleSubmit(
    context: TurnContext,
    taskModuleRequest: TaskModuleRequest
  ) {
    const conversationData: SaveData = await this.conversationDataAccessor.get(
      context,
      DEFAULTS
    )
    // Create new object to save task details.
    return this.taskManager.handleSubmit(
      taskModuleRequest,
      context,
      conversationData
    )
  }

  async run(context: TurnContext): Promise<void> {
    await super.run(context)
    await this.conversationState.saveChanges(context, false)
  }
}
