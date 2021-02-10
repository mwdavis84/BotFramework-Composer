// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';

import { emptyTemplateBodyRegex, activityTemplateType } from '../components/lg/constants';

import { getStructuredResponseFromTemplate } from './structuredResponse';

/**
 * Validates an lg template to check if it includes an structured response.
 * @param lgTemplate LgTemplate to validate.
 */
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
