// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import formatMessage from 'format-message';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import React from 'react';

import { extractTemplateNameFromExpression } from '../../../utils/structuredResponse';
import { CommonModalityEditorProps, InputHintStructuredResponseItem, SpeechStructuredResponseItem } from '../types';

import { ModalityEditorContainer } from './ModalityEditorContainer';
import { StringArrayEditor } from './StringArrayEditor';

const getInitialTemplateId = (response: SpeechStructuredResponseItem): string | undefined => {
  if (response?.value[0]) {
    return extractTemplateNameFromExpression(response.value[0]);
  }
};

const getInitialItems = (response: SpeechStructuredResponseItem, lgTemplates?: readonly LgTemplate[]): string[] => {
  const templateId = getInitialTemplateId(response);
  const template = lgTemplates?.find(({ name }) => name === templateId);
  return response?.value && template?.body
    ? template?.body?.replace(/- /g, '').split('\n') || []
    : response?.value || [];
};

type Props = CommonModalityEditorProps & {
  response: SpeechStructuredResponseItem;
  inputHint?: InputHintStructuredResponseItem['value'] | 'none';
};

const SpeechModalityEditor = React.memo(
  ({
    response,
    removeModalityDisabled: disableRemoveModality,
    lgOption,
    lgTemplates,
    memoryVariables,
    inputHint = 'none',
    onInputHintChange,
    onTemplateChange,
    onRemoveModality,
    onRemoveTemplate,
    onUpdateResponseTemplate,
  }: Props) => {
    const [templateId, setTemplateId] = React.useState(getInitialTemplateId(response));
    const [items, setItems] = React.useState<string[]>(getInitialItems(response, lgTemplates));

    const handleChange = React.useCallback(
      (newItems: string[]) => {
        setItems(newItems);
        const id = templateId || `${lgOption?.templateId}_speech`;
        if (!newItems.length) {
          setTemplateId(id);
          onUpdateResponseTemplate({ Speak: { kind: 'Speak', value: [], valueType: 'direct' } });
          onRemoveTemplate(id);
        } else if (newItems.length === 1 && lgOption?.templateId) {
          onUpdateResponseTemplate({ Speak: { kind: 'Speak', value: [newItems[0]], valueType: 'direct' } });
          onTemplateChange(id, '');
        } else {
          setTemplateId(id);
          onUpdateResponseTemplate({ Speak: { kind: 'Speak', value: [`\${${id}()}`], valueType: 'template' } });
          onTemplateChange(id, newItems.map((item) => `- ${item}`).join('\n'));
        }
      },
      [lgOption, setItems, templateId, onTemplateChange, onUpdateResponseTemplate]
    );

    const inputHintOptions = React.useMemo<IDropdownOption[]>(
      () => [
        {
          key: 'none',
          text: formatMessage('None'),
          selected: inputHint === 'none',
        },
        {
          key: 'accepting',
          text: formatMessage('Accepting'),
          selected: inputHint === 'accepting',
        },
        {
          key: 'ignoring',
          text: formatMessage('Ignoring'),
          selected: inputHint === 'ignoring',
        },
        {
          key: 'expecting',
          text: formatMessage('Expecting'),
          selected: inputHint === 'expecting',
        },
      ],
      [inputHint]
    );

    const handleInputHintChange = React.useCallback(
      (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        if (option) {
          onInputHintChange?.(option.key as string);
        }
      },
      [onInputHintChange]
    );

    return (
      <ModalityEditorContainer
        contentDescription="One of the variations added below will be selected at random by the LG library."
        contentTitle={formatMessage('Response Variations')}
        disableRemoveModality={disableRemoveModality}
        dropdownOptions={inputHintOptions}
        dropdownPrefix={formatMessage('Input hint: ')}
        modalityTitle={formatMessage('Suggested Actions')}
        modalityType="Speak"
        removeModalityOptionText={formatMessage('Remove all speech responses')}
        onDropdownChange={handleInputHintChange}
        onRemoveModality={onRemoveModality}
      >
        <StringArrayEditor
          isSpeech
          items={items}
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          onChange={handleChange}
        />
      </ModalityEditorContainer>
    );
  }
);

export { SpeechModalityEditor };
