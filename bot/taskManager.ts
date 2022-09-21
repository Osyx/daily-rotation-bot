import { TaskModuleResponse, TurnContext } from "botbuilder"
import schedule from "node-schedule"
import { TaskModuleResponseFactory } from "./TaskModuleResponseFactory"
import { DailyRotationBot } from "./dailyRotationBot"

type TaskInfo = {
  url?: string
  height?: number
  width?: number
  title?: string
}

type TaskDetail = {
  title: string
  dateTime: number
  description: string
  selectedDays: string
}

type SaveData = {
  chosenName: string
  chosenIndex: number
  registeredUserIds: string[]
}

export class TaskManager {
  taskDetails = {}
  conversationReferences = {}

  public async handleFetch(
    taskModuleRequest: any
  ): Promise<TaskModuleResponse> {
    const cardTaskFetchId = taskModuleRequest.data.id
    const taskInfo: TaskInfo = {}

    if (cardTaskFetchId === "schedule") {
      taskInfo.url = "/scheduleTask"
      taskInfo.height = 350
      taskInfo.width = 350
      taskInfo.title = "Schedule rotation"
    }

    return TaskModuleResponseFactory.toTaskModuleResponse(taskInfo)
  }

  public async handleSubmit(
    taskModuleRequest: any,
    context: TurnContext,
    conversationData: SaveData
  ) {
    const taskDetails: TaskDetail = {
      title: taskModuleRequest.data.title,
      dateTime: taskModuleRequest.data.dateTime,
      description: taskModuleRequest.data.description,
      selectedDays: taskModuleRequest.data.selectedDays
    }

    this.saveTaskDetails(taskDetails)
    await context.sendActivity(
      "Task submitted successfully, you will get a recurring reminder for the task at a scheduled time"
    )

    const currentUser = context.activity.from.id
    this.conversationReferences[currentUser] =
      TurnContext.getConversationReference(context.activity)
    const adapter = context.adapter

    const year = taskModuleRequest.data.dateTime.substring(0, 4)
    const month = taskModuleRequest.data.dateTime.substring(5, 7)
    const day = taskModuleRequest.data.dateTime.substring(8, 10)
    const hour = taskModuleRequest.data.dateTime.substring(11, 13)
    const min = taskModuleRequest.data.dateTime.substring(14, 16)
    const days = taskModuleRequest.data.selectedDays.toString()
    new Date(year, month - 1, day, hour, min)
    const cronExpression = min + " " + hour + " * * " + days

    const job = schedule.scheduleJob(cronExpression, async function () {
      await adapter.continueConversation(
        this.conversationReferences[currentUser],
        async (turnContext: TurnContext) => {
          const chosenOutput = await DailyRotationBot.handleChoose(
            context,
            conversationData,
            { chosen: conversationData.chosenName }
          )
          await turnContext.sendActivity(chosenOutput.cardActivity)
          if (chosenOutput.pingActivity !== undefined) {
            await context.sendActivity(chosenOutput.pingActivity)
          }
        }
      )
    })

    return null
  }

  // This method is used to save task details.
  saveTaskDetails(taskDetails: TaskDetail) {
    this.taskDetails["taskDetails"] = taskDetails
  }
}
