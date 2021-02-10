// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@botframework-composer/types';

const activityTemplateType = 'Activity';
const emptyTemplateBodyRegex = /^$|-(\s)?/;
const subTemplateNameRegex = /\${(.*)}/;

const acceptedInputHintValues = ['expecting', 'ignoring', 'accepting'] as const;
const acceptedAttachmentLayout = ['carousel', 'list'] as const;

const structuredResponseKeys = [
  'Text',
  'Speak',
  'Attachments',
  'AttachmentLayout',
  'InputHint',
  'SuggestedActions',
] as const;

type TextStructuredResponse = { kind: 'Text'; value: string[]; valueType: 'template' | 'direct' };
type SpeakStructuredResponse = { kind: 'Speak'; value: string[]; valueType: 'template' | 'direct' };
type AttachmentsStructuredResponse = { kind: 'Attachments'; value: string[]; valueType: 'template' | 'direct' };
type AttachmentLayoutStructuredResponse = { kind: 'AttachmentLayout'; value: typeof acceptedAttachmentLayout[number] };
type InputHintStructuredResponse = { kind: 'InputHint'; value: typeof acceptedInputHintValues[number] };
type SuggestedActionsStructuredResponse = { kind: 'SuggestedActions'; value: string[] };

type StructuredResponse =
  | TextStructuredResponse
  | SpeakStructuredResponse
  | SuggestedActionsStructuredResponse
  | InputHintStructuredResponse
  | AttachmentLayoutStructuredResponse
  | AttachmentsStructuredResponse;

const getStructuredResponseHelper = (value: unknown, kind: 'Text' | 'Speak' | 'Attachments') => {
  if (typeof value === 'string') {
    const valueAsString = value as string;
    const valueType = subTemplateNameRegex.test(valueAsString) ? 'template' : 'direct';

    return { kind, value: [valueAsString], valueType };
  } else if (Array.isArray(value)) {
    const valueAsArray = value as string[];

    return { kind, value: valueAsArray, valueType: 'direct' };
  }

  return undefined;
};

const getStructuredResponseByKind = (
  template: LgTemplate,
  kind: StructuredResponse['kind']
): StructuredResponse | undefined => {
  const value = template.properties?.[kind];
  if (value === undefined) {
    return undefined;
  }

  switch (kind) {
    case 'Text':
      return getStructuredResponseHelper(value, 'Text') as TextStructuredResponse;
    case 'Speak':
      return getStructuredResponseHelper(value, 'Speak') as SpeakStructuredResponse;
    case 'Attachments':
      return getStructuredResponseHelper(value, 'Attachments') as AttachmentsStructuredResponse;
    case 'SuggestedActions': {
      if (Array.isArray(value)) {
        const responseValue = value as string[];
        return { kind: 'SuggestedActions', value: responseValue } as SuggestedActionsStructuredResponse;
      }
      break;
    }
    case 'AttachmentLayout':
      if (acceptedAttachmentLayout.includes(value as typeof acceptedAttachmentLayout[number])) {
        return {
          kind: 'AttachmentLayout',
          value: value as typeof acceptedAttachmentLayout[number],
        } as AttachmentLayoutStructuredResponse;
      }
      break;
    case 'InputHint':
      if (acceptedInputHintValues.includes(value as typeof acceptedInputHintValues[number])) {
        return {
          kind: 'InputHint',
          value: value as typeof acceptedInputHintValues[number],
        } as InputHintStructuredResponse;
      }
      break;
  }

  return undefined;
};

/**
 * Converts template properties to structured response.
 * @param lgTemplate LgTemplate to convert.
 */
export const getStructuredResponseFromTemplate = (lgTemplate: LgTemplate) => {
  if (!lgTemplate.body || emptyTemplateBodyRegex.test(lgTemplate.body)) {
    return undefined;
  }

  if (lgTemplate.properties?.$type !== activityTemplateType) {
    return undefined;
  }

  const structuredResponse = structuredResponseKeys.reduce((response, key) => {
    const value = getStructuredResponseByKind(lgTemplate, key);
    if (value !== undefined) {
      response[key] = value;
    }

    return response;
  }, {});

  return Object.keys(structuredResponse).length ? structuredResponse : undefined;
};

export const validateStructuredResponse = (lgTemplate: LgTemplate) => {
  // If empty template return true
  if (!lgTemplate.body || emptyTemplateBodyRegex.test(lgTemplate.body)) {
    return true;
  }

  // If not of type Activity, return false
  if (lgTemplate.properties?.$type !== activityTemplateType) {
    return false;
  }

  return !!getStructuredResponseFromTemplate(lgTemplate);
};
