/**
 * Returns the contextual add-on data that should be rendered for
 * the current e-mail thread. This function satisfies the requirements of
 * an 'onTriggerFunction' and is specified in the add-on's manifest.
 *
 * @param {Object} event Event containing the message ID and other context.
 * @returns {Card}
 */
function getContextualAddOn(event) {
  // Use the GmailApp service to fetch the current email message using the event parameter
  var message = getCurrentMessage(event);
  var subject = message.getSubject();
  var snippet = message.getPlainBody();

  const date = new Date().toISOString();

  const emailPrompt =
    "You have a tool (function) that can help you create to-dos with deadlines. Extract the todos from the following email and just call the function with the extracted data. If there's no deadline, just return the action item name. If the deadline is ASAP, use the current UTC date and time. Today's date is " +
    date +
    "\n" +
    snippet +
    "  You are an assistant that only responds in JSON. Do not write normal text. You will return output in the following format action items : an array of objects. Each object has an action item name, and a due deadline in UTC date and time. Under no circumstances should you return normal text. Only respond in JSON object with action items array if any, and an empty array if none;";

  var jsonResponse = getJson(emailPrompt);
  const structured_response = JSON.parse(jsonResponse.getContentText());
  const choices = structured_response["choices"][0];
  const assistant_message = choices["message"];
  const res_content = assistant_message["tool_calls"][0];
  const res_arguments = res_content["function"]["arguments"];
  try {
    var action_items = JSON.parse(res_arguments);

    const todo_array = action_items["action_items"];

    const card = buildTasksCard(todo_array);

    return card.build();
  } catch (e) {
    return null;
  }
}

/**
 * This function shows the tasks on a card and returns the card.
 *
 * @param {todo_array} an array containing lists!
 * @returns {card}
 */
function buildTasksCard(todo_array) {
  // Create a new card builder
  var card = CardService.newCardBuilder();
  card.setHeader(
    CardService.newCardHeader().setTitle("Actionable Items from this email")
  );
  var section = CardService.newCardSection();
  // if nothing in array show nothing and return
  if (todo_array.length <= 0) {
    var addTaskButton = CardService.newTextButton()
      .setText("Add a task")
      .setOnClickAction(CardService.newAction().setFunctionName("addNewTask"));

    var addTaskLabel = CardService.newDecoratedText()
      .setText("No actionable items")
      .setButton(addTaskButton);

    section.addWidget(addTaskLabel);
    card.addSection(section);

    return card;
  }

  todo_array.forEach(function (value, index) {
    const title_field_name = "title" + index;
    // Create a text input widget with default text
    var textInput = CardService.newTextInput()
      .setFieldName(title_field_name)
      .setTitle("Enter task")
      .setValue(value["action_item_name"]); // Set default text here

    // Create a date/time picker widget
    const title_date_time = "dateTimePickerField" + index;
    var deadlineDate = new Date(value["deadline"]).getTime();
    var dateTimePicker = CardService.newDateTimePicker()
      .setFieldName(title_date_time)
      .setTitle("Select a date and time")
      .setValueInMsSinceEpoch(deadlineDate); // Set default date/time as the current date/time

    var deleteTaskButton = CardService.newTextButton()
      .setText("Delete")
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName("deleteTask")
          .setParameters({ i: String(index) })
      );

    var deleteTask = CardService.newDecoratedText()
      .setText("Remove this item")
      .setButton(deleteTaskButton);

    var taskDivider = CardService.newDivider();

    // add widget to the section.
    section.addWidget(textInput);
    section.addWidget(dateTimePicker);
    section.addWidget(deleteTask);
    section.addWidget(taskDivider);
  });
  card.addSection(section);

  if (todo_array.length > 0) {
    var buttonFixedFooter = CardService.newFixedFooter();

    buttonFixedFooter
      .setPrimaryButton(
        CardService.newTextButton()
          .setText("Add To Calendar")
          .setOnClickAction(
            CardService.newAction().setFunctionName("addToCalendar")
          )
      )
      .setSecondaryButton(
        CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(
            CardService.newAction().setFunctionName("cancelOperation")
          )
      );

    card.setFixedFooter(buttonFixedFooter);
  }

  // Add the email details section to the card

  // Return the built card
  return card;
}

/**
 * This function is responsible for creating a task on the UI
 * @param {event_object} event object, see
 *       https://developers.google.com/apps-script/add-ons/concepts/event-objects#common_event_object
 * @returns {card_response}
 */
function addNewTask(e) {
  var date = new Date();

  // get tasks on form
  var tasks = parseFormInputIntoSimplerEvents(e.commonEventObject.formInputs);

  // if there are no tasks, create card and update it
  if (tasks.length <= 0) {
    tasks = [{ action_item_name: "", deadline: date.toISOString() }];
  } else {
    tasks.push({ action_item_name: "", deadline: date.toISOString() });
  }

  const card = buildTasksCard(tasks);

  // // update the card.
  var nav = CardService.newNavigation().updateCard(card.build());

  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

/**
 * This function is responsible for deleting a task on the UI
 * @param {event_object} event object, see
 *       https://developers.google.com/apps-script/add-ons/concepts/event-objects#common_event_object
 * @returns {card_response}
 */
function deleteTask(e) {
  var tasks = parseFormInputIntoSimplerEvents(e.commonEventObject.formInputs);

  // get the index of task to be deleted.
  var index = e.commonEventObject.parameters.i;

  // parse index
  var taskIndex = parseInt(index, 10);

  // remove it from list.
  tasks.splice(taskIndex, 1);

  const card = buildTasksCard(tasks);

  // update the card.
  var nav = CardService.newNavigation().updateCard(card.build());

  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

function createDummyTasks() {
  const tasks = [
    {
      action_item_name: "Register for on-campus housing for commencement ",
      deadline: "2024-03-23T19:25:00.000Z",
    },
    {
      action_item_name: "Order your cap and gown",
      deadline: "2024-03-23T19:25:00.000Z",
    },
  ];

  return tasks;
}

// This function will be called when the 'Add To Calendar' button is clicked
/**
 * This function adds items to calendar when clicked.
 * @param {event_object} event object, see
 *       https://developers.google.com/apps-script/add-ons/concepts/event-objects#common_event_object
 * @returns {card_response}
 */
function addToCalendar(e) {
  // Assuming 'e' contains the necessary event data
  var actionItems = parseFormInputIntoSimplerEvents(
    e.commonEventObject.formInputs
  );

  var calendarInviteCount = 0;

  // Loop through each action item to create a calendar event
  for (var i = 0; i < actionItems.length; i++) {
    var item = actionItems[i];

    var event_title = item["action_item_name"];
    var event_time = item["deadline"];
    var eventId = createHourLongCalendarEvent(event_title, event_time);

    if (eventId) {
      // If an event was successfully created
      calendarInviteCount++; // Increment the successful invite counter
    }
  }

  // Create a message about the number of calendar items added
  var calendarInviteResponse =
    calendarInviteCount + " items added to Calendar.";

  return CardService.newActionResponseBuilder()
    .setNotification(
      CardService.newNotification().setText(calendarInviteResponse)
    )
    .build();
}

/**
 * Function that parses google's weird Json event common object into a streamlined array
 *
 * @params google's form inputs :
 * @returns array of action item objects [{action_item_name : name, deadline : deadline.}]
 *
 */
function parseFormInputIntoSimplerEvents(formInputs) {
  var actionItems = [];
  if (!formInputs) {
    return actionItems;
  }

  for (var i = 0; ; i++) {
    var titleKey = "title" + i;
    var dateTimeKey = "dateTimePickerField" + i;

    if (formInputs[titleKey] && formInputs[dateTimeKey]) {
      var actionItemName = formInputs[titleKey][""]["stringInputs"]["value"][0];
      // convert to local time with offset cuz everything else is crazy
      var offset = new Date().getTimezoneOffset() * 60000;
      var msSinceEpoch =
        formInputs[dateTimeKey][""]["dateTimeInput"]["msSinceEpoch"] + offset;
      // convert to local time with offset cuz everything else is crazy
      var deadline = new Date(msSinceEpoch);

      actionItems.push({
        action_item_name: actionItemName,
        deadline: deadline.toISOString(), // Format the date as ISO string
      });
    } else {
      // If we can't find a matching title or dateTime field, we assume we've processed all items
      break;
    }
  }

  return actionItems;
}

/**
 * Adds a one hour long calendar event to the caller's default calendar
 *
 * @param {string} eventName the name of the event
 * @param {string} startDatetime the starting date time of the event
 */
function createHourLongCalendarEvent(title, startDatetime) {
  const eventName = typeof title === "string" ? title : "My event";
  const startDate = new Date(startDatetime);

  const endTime = addHoursToDate(startDate, 1);

  try {
    // Create the event using CalendarApp
    var event = CalendarApp.getDefaultCalendar().createEvent(
      eventName,
      startDate,
      endTime
    );

    return event.getId(); // Return the ID of the created event
  } catch (e) {
    Logger.log(e.message);
    return null; // Return null if an error occurs
  }
}

/**
 * Returns a new date with x hours added to it.
 * Meant as a helper function
 *
 * @param {Date} startingDate the starting date to add hours to
 * @param {number} hours number of hours to add
 * @return {Date} the new date
 */
function addHoursToDate(startingDate, hours) {
  var newDate = new Date(startingDate.getTime());
  newDate.setHours(newDate.getHours() + hours);
  return newDate;
}

function getCurrentMessage(event) {
  var accessToken = event.messageMetadata.accessToken;
  var messageId = event.messageMetadata.messageId;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  return GmailApp.getMessageById(messageId);
}

tools = [
  {
    type: "function",
    function: {
      name: "create_todos",
      description:
        "Creates a list of to-dos with specified deadlines in UTC format",
      parameters: {
        type: "object",
        properties: {
          action_items: {
            available_action_items: "boolean",
            type: "array",
            items: {
              type: "object",
              properties: {
                action_item_name: {
                  type: "string",
                  description: "The name of the action item.",
                },
                deadline: {
                  type: "string",
                  format: "date-time",
                  description:
                    "The UTC deadline for the action item, formatted as 'YYYY-MM-DDTHH:MM:SSZ'. For immediate or ASAP tasks, use the current UTC date and time.",
                },
              },
              required: ["action_item_name", "deadline"],
            },
            description:
              "A list of action items each with a name and a UTC formatted deadline.",
          },
        },
        required: ["action_items"],
      },
    },
  },
];

var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_KEY");

function getJson(prompt) {
  var apiUrl = "https://api.openai.com/v1/chat/completions";
  var payload = JSON.stringify({
    model: "gpt-4",
    messages: [
      {
        role: "user",
        content: prompt,
      },
    ],
    temperature: 0.001,
    max_tokens: 256,
    top_p: 1,
    tools: tools,
    frequency_penalty: 0,
    presence_penalty: 0,
  });

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey,
    },
    payload: payload,
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(apiUrl, options);

  return response;
}
