// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import formatMessage from 'format-message';
import React, { useCallback, useState } from 'react';

import { extractTemplateNameFromExpression } from '../../../utils/structuredResponse';
import { CommonModalityEditorProps, TextStructuredResponseItem } from '../types';

import { ModalityEditorContainer } from './ModalityEditorContainer';
import { StringArrayEditor } from './StringArrayEditor';

const getInitialTemplateId = (response: TextStructuredResponseItem): string | undefined => {
  if (response?.value) {
    return extractTemplateNameFromExpression(Array.isArray(response.value) ? response.value[0] : response.value);
  }
};

const getInitialItems = (response: TextStructuredResponseItem, lgTemplates?: readonly LgTemplate[]): string[] => {
  const templateId = getInitialTemplateId(response);
  const template = lgTemplates?.find(({ name }) => name === templateId);
  return response?.value && template?.body
    ? template?.body?.replace(/- /g, '').split('\n') || []
    : response?.value || [];
};

type Props = CommonModalityEditorProps & { response: TextStructuredResponseItem };

const TextModalityEditor = React.memo(
  ({
    response,
    removeModalityDisabled: disableRemoveModality,
    lgOption,
    lgTemplates,
    memoryVariables,
    onTemplateChange,
    onRemoveModality,
    onUpdateResponseTemplate,
  }: Props) => {
    const [templateId, setTemplateId] = React.useState(getInitialTemplateId(response));
    const [items, setItems] = useState<string[]>(getInitialItems(response, lgTemplates));

    const handleChange = useCallback(
      (newItems: string[]) => {
        setItems(newItems);
        if (newItems.length === 1 && lgOption?.templateId) {
          onUpdateResponseTemplate({ Text: { kind: 'Text', value: [newItems[0]], valueType: 'direct' } });
        } else {
          const id = templateId || `${lgOption?.templateId}_text`;
          setTemplateId(id);
          onUpdateResponseTemplate({ Text: { kind: 'Text', value: [`\${${id}()}`], valueType: 'template' } });
          onTemplateChange(id, newItems.map((item) => `- ${item}`).join('\n'));
        }
      },
      [lgOption, setItems, templateId, onTemplateChange, onUpdateResponseTemplate]
    );

    return (
      <ModalityEditorContainer
        contentDescription={formatMessage(
          'One of the variations added below will be selected at random by the LG library.'
        )}
        contentTitle={formatMessage('Response Variations')}
        disableRemoveModality={disableRemoveModality}
        modalityTitle={formatMessage('Text')}
        modalityType="Text"
        removeModalityOptionText={formatMessage('Remove all text responses')}
        onRemoveModality={onRemoveModality}
      >
        <StringArrayEditor
          items={items}
          lgOption={lgOption}
          lgTemplates={lgTemplates}
          memoryVariables={memoryVariables}
          selectedKey="text"
          onChange={handleChange}
        />
      </ModalityEditorContainer>
    );
  }
);

export { TextModalityEditor };
