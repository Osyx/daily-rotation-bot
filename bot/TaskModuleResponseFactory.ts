// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
import { TaskModuleResponse } from "botbuilder"

type TaskInfo = {
  url?: string
  height?: number
  width?: number
  title?: string
}

export class TaskModuleResponseFactory {
  static createResponse(taskModuleInfoOrString: any): TaskModuleResponse {
    if (typeof taskModuleInfoOrString === "string") {
      return {
        task: {
          type: "message",
          value: taskModuleInfoOrString
        }
      }
    }

    return {
      task: {
        type: "continue",
        value: taskModuleInfoOrString
      }
    }
  }

  static toTaskModuleResponse(taskInfo: TaskInfo): TaskModuleResponse {
    return TaskModuleResponseFactory.createResponse(taskInfo)
  }
}
