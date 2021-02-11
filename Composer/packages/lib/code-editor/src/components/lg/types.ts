// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';

import { LGOption } from '../../utils';

export type TemplateRefPayload = {
  kind: 'templateRef';
  data: { templates: readonly LgTemplate[]; onSelectTemplate: (templateString: string) => void };
};

export type PropertyItem = {
  id: string;
  children: PropertyItem[];
};

export type PropertyRefPayload = {
  kind: 'propertyRef';
  data: { properties: readonly string[]; onSelectProperty: (property: string) => void };
};

export type FunctionRefPayload = {
  kind: 'functionRef';
  data: {
    functions: readonly { key: string; name: string; children: string[] }[];
    onSelectFunction: (functionString: string) => void;
  };
};

export type ToolbarButtonPayload = TemplateRefPayload | PropertyRefPayload | FunctionRefPayload;

export type LgLanguageContext =
  | 'expression'
  | 'singleQuote'
  | 'doubleQuote'
  | 'comment'
  | 'templateBody'
  | 'templateName'
  | 'root';

export type PartialStructuredResponse = Partial<Record<StructuredResponseItem['kind'], StructuredResponseItem>>;

export type CommonModalityEditorProps = {
  response?: PartialStructuredResponse;
  removeModalityDisabled: boolean;
  lgOption?: LGOption;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  onAttachmentLayoutChange?: (layout: string) => void;
  onInputHintChange?: (inputHint: string) => void;
  onTemplateChange: (templateId: string, body?: string) => void;
  onRemoveModality: (modality: ModalityType) => void;
  onUpdateResponseTemplate: (response: PartialStructuredResponse) => void;
};

/**
 * Structured response types.
 */
export const acceptedInputHintValues = ['expecting', 'ignoring', 'accepting'] as const;
export const acceptedAttachmentLayout = ['carousel', 'list'] as const;

export const modalityType = ['Text', 'Speak', 'Attachments', 'SuggestedActions'] as const;
export const structuredResponseKeys = [...modalityType, 'AttachmentLayout', 'InputHint'] as const;

export type ModalityType = typeof modalityType[number];

export type TextStructuredResponseItem = { kind: 'Text'; value: string[]; valueType: 'template' | 'direct' };
export type SpeechStructuredResponseItem = { kind: 'Speak'; value: string[]; valueType: 'template' | 'direct' };
export type AttachmentsStructuredResponseItem = {
  kind: 'Attachments';
  value: string[];
  valueType: 'template' | 'direct';
};
export type AttachmentLayoutStructuredResponseItem = {
  kind: 'AttachmentLayout';
  value: typeof acceptedAttachmentLayout[number];
};
export type InputHintStructuredResponseItem = { kind: 'InputHint'; value: typeof acceptedInputHintValues[number] };
export type SuggestedActionsStructuredResponseItem = { kind: 'SuggestedActions'; value: string[] };

export type StructuredResponseItem =
  | TextStructuredResponseItem
  | SpeechStructuredResponseItem
  | SuggestedActionsStructuredResponseItem
  | InputHintStructuredResponseItem
  | AttachmentLayoutStructuredResponseItem
  | AttachmentsStructuredResponseItem;
