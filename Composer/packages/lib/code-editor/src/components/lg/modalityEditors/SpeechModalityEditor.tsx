// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import formatMessage from 'format-message';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import React from 'react';

import { CommonModalityEditorProps, SpeakStructuredResponseItem } from '../types';

import { ModalityEditorContainer } from './ModalityEditorContainer';
import { StringArrayEditor } from './StringArrayEditor';

type Props = CommonModalityEditorProps & { response: SpeakStructuredResponseItem };

const SpeechModalityEditor = React.memo(
  ({
    removeModalityDisabled: disableRemoveModality,
    lgOption,
    lgTemplates,
    memoryVariables,
    onInputHintChange,
    onTemplateChange,
    onRemoveModality,
  }: Props) => {
    const [items, setItems] = React.useState<string[]>([]);
    const [templateId] = React.useState<string>(`${lgOption?.templateId}_speak}`);

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
          selected: true,
        },
        {
          key: 'acceptingInput',
          text: formatMessage('Accepting'),
        },
        {
          key: 'ignoringInput',
          text: formatMessage('Ignoring'),
        },
        {
          key: 'expectingInput',
          text: formatMessage('Expecting'),
        },
      ],
      []
    );

    const handleInputHintChange = React.useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
      if (option) {
        onInputHintChange?.(option.key as string);
      }
    }, []);

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
          items={items}
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          selectedKey="speak"
          onChange={handleChange}
        />
      </ModalityEditorContainer>
    );
  }
);

export { SpeechModalityEditor };
