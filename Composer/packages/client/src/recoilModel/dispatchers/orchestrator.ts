import { ITrigger, SDKKinds } from '@bfc/shared';
import { camelCase, cloneDeep } from 'lodash';
import { CallbackInterface, useRecoilCallback } from 'recoil';
import { dialogsSelectorFamily, dialogState, dialogsWithLuProviderSelectorFamily } from '..';
import { insert } from '../../utils/dialogUtil';

const generateId = () => {
  const arr = crypto.getRandomValues(new Uint32Array(1));
  return `${arr[0]}`;
};

const automaticSkillIntentTrigger = (skillName: string) => {
  let camelCasedSkillName = camelCase(skillName);
  return {
    $kind: 'Microsoft.OnIntent',
    $designer: {
      id: generateId(),
      name: skillName,
    },
    intent: skillName,
    actions: [
      {
        $kind: 'Microsoft.BeginSkill',
        $designer: {
          id: generateId(),
        },
        activityProcessed: true,
        botId: '=settings.MicrosoftAppId',
        skillHostEndpoint: '=settings.skillHostEndpoint',
        connectionName: '=settings.connectionName',
        allowInterruptions: true,
        skillEndpoint: `=settings.skill['${camelCasedSkillName}'].endpointUrl`,
        skillAppId: `=settings.skill['${camelCasedSkillName}'].msAppId`,
      },
    ],
  };
};

export const orchestratorDispatcher = () => {
  const createAutomaticTrigger = useRecoilCallback(
    (callbackHelpers: CallbackInterface) => async (rootBotProjectId: string, childBotProjectId: string) => {
      const { set, snapshot } = callbackHelpers;
      const rootDialogs = await snapshot.getPromise(dialogsWithLuProviderSelectorFamily(rootBotProjectId));

      try {
        const rootDialog = rootDialogs.filter((d) => d.isRoot)?.[0];

        if (rootDialog?.luProvider === SDKKinds.OrchestratorRecognizer) {
          //const dispatcher = await snapshot.getPromise(dispatcherState);

          const dialogs = await snapshot.getPromise(dialogsSelectorFamily(childBotProjectId));
          let skillName = dialogs.filter((d) => d.isRoot).map((d) => d.luFile)?.[0];

          //const triggerCopy = Object.assign([], rootDialog.content.triggers);
          //triggerCopy.push(automaticSkillIntentTrigger(skillName, skillName));

          const dialogCopy = cloneDeep(rootDialog);
          insert(dialogCopy.content, 'triggers', undefined, automaticSkillIntentTrigger(skillName));
          insert(dialogCopy, 'triggers', undefined, {
            displayName: skillName,
            isIntent: true,
            type: SDKKinds.OnIntent,
            content: automaticSkillIntentTrigger(skillName),
          } as ITrigger);

          //   const newRootDialog: DialogInfo = {
          //     ...rootDialog,
          //     content: {
          //         ...content,
          //         triggers: triggerCopy
          //     }
          //   };

          set(dialogState({ projectId: rootBotProjectId, dialogId: rootDialog.id }), dialogCopy);
        }
      } catch (err) {
        //set(botStatusState(projectId), BotStatus.failed);
        // set(botRuntimeErrorState(projectId), {
        //   title: Text.LUISDEPLOYFAILURE,
        //   message: err.response?.data?.message || err.message,
        // });
        console.log(err);
      }
    }
  );

  return {
    createAutomaticTrigger,
  };
};
