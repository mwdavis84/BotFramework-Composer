// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
/* eslint-disable react-hooks/rules-of-hooks */

import { useRecoilCallback, CallbackInterface } from 'recoil';
import { DialogInfo, ILuisConfig, IQnAConfig, ITrigger, LuFile, LUISLocales, SDKKinds } from '@bfc/shared';
import formatMessage from 'format-message';
import difference from 'lodash/difference';

import * as luUtil from '../../utils/luUtil';
import { Text, BotStatus } from '../../constants';
import httpClient from '../../utils/httpUtil';
import luFileStatusStorage from '../../utils/luFileStatusStorage';
import qnaFileStatusStorage from '../../utils/qnaFileStatusStorage';
import { LuProviderType } from './../../../types/src/indexers';
import {
  luFilesState,
  qnaFilesState,
  botStatusState,
  botRuntimeErrorState,
  settingsState,
  dialogState,
} from '../atoms';
import {
  dialogsSelectorFamily,
  dialogsWithLuProviderSelectorFamily,
  localBotsWithoutErrorsSelector,
} from '../selectors';
import { getReferredQnaFiles } from '../../utils/qnaUtil';

import { addNotificationInternal, createNotification } from './notification';
import { getLuProvider } from '../../utils/dialogUtil';

const checkEmptyQuestionOrAnswerInQnAFile = (sections) => {
  return sections.some((s) => !s.Answer || s.Questions.some((q) => !q.content));
};

const setLuisBuildNotification = (callbackHelpers: CallbackInterface, unsupportedLocales: string[]) => {
  if (!unsupportedLocales.length) return;
  const notification = createNotification({
    title: formatMessage('Luis build warning'),
    description: formatMessage('locale "{locale}" is not supported by LUIS', { locale: unsupportedLocales.join(' ') }),
    type: 'warning',
    retentionTime: 5000,
  });
  addNotificationInternal(callbackHelpers, notification);
};

const crossProjectBuild = async (
  projectId: string,
  rootBotDialogs: DialogInfo,
  rootBotLuFiles: LuFile[],
  skillLUFiles: LuFile[]
) => {
  //supports orchestrator only
  if (rootBotDialogs.luProvider !== SDKKinds.OrchestratorRecognizer) {
    return;
  }
  //const parentLU = rootBotLuFiles.filter((lu)=>lu.id === rootBotDialogs.id);

  return await httpClient.post(`/projects/${projectId}/crossbuild`, {
    parentLU: rootBotLuFiles,
    luFilesToMerge: skillLUFiles,
  });
};

export const builderDispatcher = () => {
  const build = useRecoilCallback(
    (callbackHelpers: CallbackInterface) => async (
      projectId: string,
      luisConfig: ILuisConfig,
      qnaConfig: IQnAConfig
    ) => {
      const { set, snapshot } = callbackHelpers;
      const dialogs = await snapshot.getPromise(dialogsWithLuProviderSelectorFamily(projectId));
      const luFiles = await snapshot.getPromise(luFilesState(projectId));
      const qnaFiles = await snapshot.getPromise(qnaFilesState(projectId));
      const { languages } = await snapshot.getPromise(settingsState(projectId));
      const referredLuFiles = luUtil.checkLuisBuild(luFiles, dialogs);
      const referredQnaFiles = getReferredQnaFiles(qnaFiles, dialogs, false);
      const unsupportedLocales = difference<string>(languages, LUISLocales);
      setLuisBuildNotification(callbackHelpers, unsupportedLocales);
      const errorMsg = referredQnaFiles.reduce(
        (result, file) => {
          if (
            file.qnaSections &&
            file.qnaSections.length > 0 &&
            checkEmptyQuestionOrAnswerInQnAFile(file.qnaSections)
          ) {
            result.message = result.message + `${file.id}.qna file contains empty answer or questions`;
          }
          return result;
        },
        { title: Text.LUISDEPLOYFAILURE, message: '' }
      );
      if (errorMsg.message) {
        set(botRuntimeErrorState(projectId), errorMsg);
        set(botStatusState(projectId), BotStatus.failed);
        return;
      }
      try {
        await httpClient.post(`/projects/${projectId}/build`, {
          luisConfig,
          qnaConfig,
          projectId,
          luFiles: referredLuFiles.map((file) => ({ id: file.id, isEmpty: file.empty })),
          qnaFiles: referredQnaFiles.map((file) => ({ id: file.id, isEmpty: file.empty })),
        });
        luFileStatusStorage.publishAll(projectId);
        qnaFileStatusStorage.publishAll(projectId);
        set(botStatusState(projectId), BotStatus.published);

        //do orchestrator skill integration
        const rootDialogs = dialogs.filter((d) => d.isRoot)?.[0];

        if (rootDialogs.luProvider === SDKKinds.OrchestratorRecognizer) {
          const rootDialog = dialogs.filter((d) => d.isRoot)?.[0];
          const parentLU = luFiles.filter((lu) => lu.id.startsWith(rootDialog.luFile));
          let skillLus: LuFile[] = [];

          const botProjects = await snapshot.getPromise(localBotsWithoutErrorsSelector);
          for (var project of botProjects.filter((p) => p !== projectId)) {
            const dialogs = await snapshot.getPromise(dialogsWithLuProviderSelectorFamily(project));
            let rootDialog = dialogs.filter((d) => d.isRoot)?.[0];

            if (rootDialog.luProvider) {
              let lus = await snapshot.getPromise(luFilesState(project));
              let localSkillLu = lus.find((lu) => lu.id.startsWith(rootDialog.luFile) && !lu.empty);
              if (localSkillLu) {
                skillLus.push(localSkillLu);
              }
            }
          }
          await crossProjectBuild(projectId, rootDialog, parentLU, skillLus);
        }
      } catch (err) {
        set(botStatusState(projectId), BotStatus.failed);
        set(botRuntimeErrorState(projectId), {
          title: Text.LUISDEPLOYFAILURE,
          message: err.response?.data?.message || err.message,
        });
      }
    }
  );

  return {
    build,
  };
};
