// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import formatMessage from 'format-message';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import React from 'react';

import { CommonModalityEditorProps, InputHintStructuredResponseItem, SpeechStructuredResponseItem } from '../types';

import { ModalityEditorContainer } from './ModalityEditorContainer';
import { StringArrayEditor } from './StringArrayEditor';

type Props = CommonModalityEditorProps & {
  response: SpeechStructuredResponseItem;
  inputHint?: InputHintStructuredResponseItem['value'] | 'none';
};

const SpeechModalityEditor = React.memo(
  ({
    removeModalityDisabled: disableRemoveModality,
    lgOption,
    lgTemplates,
    memoryVariables,
    inputHint = 'none',
    onInputHintChange,
    onTemplateChange,
    onRemoveModality,
  }: Props) => {
    const [items, setItems] = React.useState<string[]>([]);
    const [templateId] = React.useState<string>(`${lgOption?.templateId}_speech}`);

    const handleChange = React.useCallback(
      (newItems: string[]) => {
        setItems(newItems);
        onTemplateChange(templateId, newItems.map((item) => `- ${item}`).join('\n'));
      },
      [setItems, templateId, onTemplateChange]
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
