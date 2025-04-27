/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
const SHOPPING_LIST = 'Shopping List';

export function onOpen(e: any) {
  SpreadsheetApp.getUi()
    .createMenu('Shopping List')
    .addItem('Create shopping list from this sheet', 'createShoppingList')
    .addToUi();
}

export function createShoppingList() {
  const shoppingList = GetOrCreateShoppingListTaskList();
  const foodColumn = GetFoodColumn().sort((a, b) => b.localeCompare(a));

  const dateTask = Tasks.newTask();
  dateTask.title = SpreadsheetApp.getActiveSheet().getName();
  Tasks.Tasks?.insert(dateTask, shoppingList.id as string);
  const dt = GetTask(dateTask.title, shoppingList.title as string);
  for (let i = 1; i < foodColumn.length; i++) {
    const currentIngredient = foodColumn[i];
    const newTask = Tasks.newTask();
    newTask.title = currentIngredient;
    Tasks.Tasks?.insert(newTask, shoppingList.id as string, {parent: dt.id});
  }
}

function GetOrCreateShoppingListTaskList(): GoogleAppsScript.Tasks.Schema.TaskList {
  try {
    return GetTaskList(SHOPPING_LIST);
  } catch {
    const mealPlanTaskList = Tasks.newTaskList();
    mealPlanTaskList.title = SHOPPING_LIST;
    Tasks.Tasklists?.insert(mealPlanTaskList);
    return GetTaskList(SHOPPING_LIST);
  }
}

function GetTaskList(taskListName: string) {
  const taskList = Tasks.Tasklists?.list().items?.find(
    taskList => taskList.title === taskListName
  );
  if (taskList === undefined)
    throw new Error(`${taskListName} could not be found`);
  return taskList;
}

function GetTask(taskName: string, taskListName: string) {
  const taskList = GetTaskList(taskListName);
  const task = Tasks.Tasks?.list(taskList.id as string).items?.find(
    task => task.title === taskName
  );
  if (task === undefined) throw new Error(`${taskName} could not be found`);
  return task;
}

function GetFoodColumn(): Array<string> {
  const range = SpreadsheetApp.getActiveSheet().getRange(1, 1, 1000, 100);
  const values = range.getValues();
  const foodColumn = values[0].findIndex((v: string) => v === 'food');
  return values.map(r => r[foodColumn]).filter(v => v !== '');
}

export function hello() {
  return 'Hello Apps Script!';
}
