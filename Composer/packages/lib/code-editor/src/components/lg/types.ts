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

export const modalityTypes = ['text', 'speak', 'attachments', 'suggestedActions'] as const;
export type ModalityType = typeof modalityTypes[number];

export type CommonModalityEditorProps = {
  removeModalityDisabled: boolean;
  lgOption?: LGOption;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  template?: LgTemplate;
  templateId: string;
  onAttachmentLayoutChange?: (layout: string) => void;
  onInputHintChange?: (inputHint: string) => void;
  onTemplateChange: (templateId: string, body?: string) => void;
  onRemoveModality: () => void;
};

/**
 * Structured response types.
 */
export const acceptedInputHintValues = ['expecting', 'ignoring', 'accepting'] as const;
export const acceptedAttachmentLayout = ['carousel', 'list'] as const;

export const structuredResponseKeys = [
  'Text',
  'Speak',
  'Attachments',
  'AttachmentLayout',
  'InputHint',
  'SuggestedActions',
] as const;

export type TextStructuredResponse = { kind: 'Text'; value: string[]; valueType: 'template' | 'direct' };
export type SpeakStructuredResponse = { kind: 'Speak'; value: string[]; valueType: 'template' | 'direct' };
export type AttachmentsStructuredResponse = { kind: 'Attachments'; value: string[]; valueType: 'template' | 'direct' };
export type AttachmentLayoutStructuredResponse = {
  kind: 'AttachmentLayout';
  value: typeof acceptedAttachmentLayout[number];
};
export type InputHintStructuredResponse = { kind: 'InputHint'; value: typeof acceptedInputHintValues[number] };
export type SuggestedActionsStructuredResponse = { kind: 'SuggestedActions'; value: string[] };

export type StructuredResponse =
  | TextStructuredResponse
  | SpeakStructuredResponse
  | SuggestedActionsStructuredResponse
  | InputHintStructuredResponse
  | AttachmentLayoutStructuredResponse
  | AttachmentsStructuredResponse;
