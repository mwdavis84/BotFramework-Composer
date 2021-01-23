// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/** @jsx jsx */
import { jsx } from '@emotion/core';
import React, { useCallback } from 'react';
import { LgEditor } from '@bfc/code-editor';
import { FieldProps, useShellApi } from '@bfc/extension-client';
import { FieldLabel, useFormData } from '@bfc/adaptive-form';
import { LgMetaData, LgTemplateRef, LgType, CodeEditorSettings } from '@bfc/shared';
import { filterTemplateDiagnostics } from '@bfc/indexers';

import { locateLgTemplatePosition } from './locateLgTemplatePosition';

const lspServerPath = '/lg-language-server';

const tryGetLgMetaDataType = (lgText: string): string | null => {
  const lgRef = LgTemplateRef.parse(lgText);
  if (lgRef === null) return null;

  const lgMetaData = LgMetaData.parse(lgRef.name);
  if (lgMetaData === null) return null;

  return lgMetaData.type;
};

const getInitialTemplate = (fieldName: string, formData?: string): string => {
  const lgText = formData || '';

  // Field content is already a ref created by composer.
  if (tryGetLgMetaDataType(lgText) === fieldName) {
    return '';
  }
  return lgText.startsWith('-') ? lgText : `- ${lgText}`;
};

const LgField: React.FC<FieldProps<string>> = (props) => {
  const { label, id, description, value, name, uiOptions, required } = props;
  const { designerId, currentDialog, lgFiles, shellApi, projectId, locale, userSettings } = useShellApi();
  const formData = useFormData();

  let lgType = name;
  const $kind = formData?.$kind;
  if ($kind) {
    lgType = new LgType($kind, name).toString();
  }

  const lgTemplateRef = LgTemplateRef.parse(value);
  const lgName = lgTemplateRef ? lgTemplateRef.name : new LgMetaData(lgType, designerId || '').toString();

  const relatedLgFile = locateLgTemplatePosition(lgFiles, lgName, locale);

  const fallbackLgFileId = `${currentDialog.lgFile}.${locale}`;
  const lgFile = relatedLgFile ?? lgFiles.find((f) => f.id === fallbackLgFileId);
  const lgFileId = lgFile?.id ?? fallbackLgFileId;

  const allTemplates = React.useMemo(
    () =>
      (lgFiles.find((lgFile) => lgFile.id === lgFileId)?.allTemplates || [])
        .filter((t) => t.name !== lgTemplateRef?.name)
        .sort(),
    [lgFileId, lgFiles]
  );

  const updateLgTemplate = useCallback(
    async (body: string) => {
      await shellApi.debouncedUpdateLgTemplate(lgFileId, lgName, body);
    },
    [lgName, lgFileId]
  );

  const template = lgFile?.templates?.find((template) => {
    return template.name === lgName;
  }) || {
    name: lgName,
    parameters: [],
    body: getInitialTemplate(name, value),
  };

  const diagnostics = lgFile ? filterTemplateDiagnostics(lgFile, template.name) : [];

  const lgOption = {
    projectId,
    fileId: lgFileId,
    templateId: lgName,
  };

  const onChange = (body: string) => {
    if (designerId) {
      if (body) {
        updateLgTemplate(body).then(() => {
          if (lgTemplateRef) {
            shellApi.commitChanges();
          }
        });
        props.onChange(new LgTemplateRef(lgName).toString());
      } else {
        shellApi.removeLgTemplate(lgFileId, lgName).then(() => {
          props.onChange();
        });
      }
    }
  };

  const handleSettingsChange = (settings: Partial<CodeEditorSettings>) => {
    shellApi.updateUserSettings({ codeEditor: settings });
  };

  return (
    <React.Fragment>
      <FieldLabel description={description} helpLink={uiOptions?.helpLink} id={id} label={label} required={required} />
      <LgEditor
        hidePlaceholder
        allTemplates={allTemplates}
        diagnostics={diagnostics}
        editorSettings={userSettings.codeEditor}
        height={225}
        languageServer={{
          path: lspServerPath,
        }}
        lgOption={lgOption}
        value={template.body}
        onChange={onChange}
        onChangeSettings={handleSettingsChange}
      />
    </React.Fragment>
  );
};

export { LgField };
