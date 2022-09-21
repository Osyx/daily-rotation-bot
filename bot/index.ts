import {
  BotFrameworkAdapter,
  ConversationState,
  MemoryStorage,
  TurnContext
} from "botbuilder"
import * as restify from "restify"
import { DailyRotationBot } from "./dailyRotationBot"

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD
})

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`)

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  )

  // Send a message to the user
  await context.sendActivity(
    `The bot encountered unhandled error:\n ${error.message}`
  )
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  )

  await conversationState.delete(context)
}

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler

// Create the bot that will handle incoming messages.
const memoryStorage = new MemoryStorage()
const conversationState = new ConversationState(memoryStorage)
const bot = new DailyRotationBot(conversationState)

// Create HTTP server.
const server = restify.createServer()
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\n${server.name} listening to ${server.url}.`)
})

// Listen for incoming requests.
server.get("/scheduleTask", (req, res, next) => {
  res.send("./views/ScheduleTask")
})

server.post("/api/messages", async (req, res) =>
  adapter.processActivity(req, res, async (context) => bot.run(context))
)
