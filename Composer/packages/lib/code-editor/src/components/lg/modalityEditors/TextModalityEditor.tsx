// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import formatMessage from 'format-message';
import React from 'react';

import { extractTemplateNameFromExpression } from '../../../utils/structuredResponse';
import { CommonModalityEditorProps, TextStructuredResponseItem } from '../types';

import { ModalityEditorContainer } from './ModalityEditorContainer';
import { StringArrayEditor } from './StringArrayEditor';

const getInitialTemplateId = (response: TextStructuredResponseItem): string | undefined => {
  if (response?.value[0]) {
    return extractTemplateNameFromExpression(response.value[0]);
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
    onRemoveTemplate,
    onUpdateResponseTemplate,
  }: Props) => {
    const [templateId, setTemplateId] = React.useState(getInitialTemplateId(response));
    const [items, setItems] = React.useState<string[]>(getInitialItems(response, lgTemplates));

    const handleChange = React.useCallback(
      (newItems: string[]) => {
        setItems(newItems);
        const id = templateId || `${lgOption?.templateId}_text`;
        if (!newItems.length) {
          setTemplateId(id);
          onUpdateResponseTemplate({ Text: { kind: 'Text', value: [], valueType: 'direct' } });
          onRemoveTemplate(id);
        } else if (newItems.length === 1 && lgOption?.templateId) {
          onUpdateResponseTemplate({ Text: { kind: 'Text', value: [newItems[0]], valueType: 'direct' } });
          onTemplateChange(id, '');
        } else {
          setTemplateId(id);
          onUpdateResponseTemplate({ Text: { kind: 'Text', value: [`\${${id}()}`], valueType: 'template' } });
          onTemplateChange(id, newItems.map((item) => `- ${item}`).join('\n'));
        }
      },
      [lgOption, setItems, templateId, onRemoveTemplate, onTemplateChange, onUpdateResponseTemplate]
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
          onChange={handleChange}
        />
      </ModalityEditorContainer>
    );
  }
);

export { TextModalityEditor };
