// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { LgTemplate } from '@bfc/shared';
import { FluentTheme } from '@uifabric/fluent-theme';
import formatMessage from 'format-message';
import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { ILinkStyles, Link } from 'office-ui-fabric-react/lib/Link';
import React, { useCallback, useEffect, useRef, useState } from 'react';

import { LGOption } from '../../../utils';
import { getCursorContextWithinLine } from '../../../utils/lgUtils';
import { jsLgToolbarMenuClassName } from '../constants';
import { LgEditorToolbar } from '../LgEditorToolbar';
import { LgSpeakModalityToolbar, SSMLTagType } from '../LgSpeakModalityToolbar';

import { StringArrayItem } from './StringArrayItem';

const submitKeys = ['Enter', 'Escape'];

const styles: { link: ILinkStyles } = {
  link: {
    root: {
      height: 32,
      paddingLeft: 13,
      fontSize: FluentTheme.fonts.small.fontSize,
      ':hover': { textDecoration: 'none' },
      ':active': { textDecoration: 'none' },
    },
  },
};

const prosodyDefaultProps = ['pitch', 'rate', 'volume'];
const breakDefaultProps = ['time', 'strength'];
const audioDefaultProps = ['src'];

const getSSMLProps = (tag: 'prosody' | 'audio' | 'break'): string => {
  let defaultProps: string[] = [];
  switch (tag) {
    case 'prosody':
      defaultProps = prosodyDefaultProps;
      break;
    case 'audio':
      defaultProps = audioDefaultProps;
      break;
    case 'break':
      defaultProps = breakDefaultProps;
  }

  return defaultProps.map((prop) => `${prop}=""`).join(' ');
};

type StringArrayEditorProps = {
  items: string[];
  selectedKey: string;
  lgTemplates?: readonly LgTemplate[];
  memoryVariables?: readonly string[];
  lgOption?: LGOption;
  onChange: (items: string[]) => void;
};

const StringArrayEditor = React.memo(
  ({ items, lgTemplates, memoryVariables, selectedKey, onChange }: StringArrayEditorProps) => {
    const containerRef = useRef<HTMLDivElement | null>(null);
    const [currentIndex, setCurrentIndex] = useState<number | null>(null);
    const [calloutTargetElement, setCalloutTargetElement] = useState<HTMLInputElement | null>(null);

    const handleChange = useCallback(
      (index: number) => (_, newValue?: string) => {
        const updatedItems = [...items];
        updatedItems[index] = newValue ?? '';
        onChange(updatedItems);
      },
      [items, onChange]
    );

    const handleFocus = useCallback(
      (index: number) => () => {
        setCurrentIndex(index);
      },
      [setCurrentIndex]
    );

    const handleRemove = useCallback(
      (index: number) => () => {
        const newItems = items.slice();
        newItems.splice(index, 1);
        onChange(newItems);
      },
      [items, onChange]
    );

    const handleClickAddVariation = useCallback(() => {
      onChange([...items, '']);
      setCurrentIndex(items.length);
    }, [items, setCurrentIndex, onChange]);

    const handleShowCallout = useCallback((targetElement: HTMLInputElement) => {
      setCalloutTargetElement(targetElement);
    }, []);

    useEffect(() => {
      const keydownHandler = (e: KeyboardEvent) => {
        if (submitKeys.includes(e.key)) {
          setCalloutTargetElement(null);
          setCurrentIndex(null);
          // Remove empty variations
          onChange(items.filter(Boolean));
        }
      };

      const focusHandler = (e: FocusEvent) => {
        if (containerRef.current?.contains(e.target as Node)) {
          return;
        }

        if (
          !e
            .composedPath()
            .filter((n) => n instanceof Element)
            .map((n) => (n as Element).className)
            .some((c) => c.indexOf(jsLgToolbarMenuClassName) !== -1)
        ) {
          setCalloutTargetElement(null);
          setCurrentIndex(null);
          // Remove empty variations
          onChange(items.filter(Boolean));
        }
      };

      document.addEventListener('keydown', keydownHandler);
      document.addEventListener('focusin', focusHandler);

      return () => {
        document.removeEventListener('keydown', keydownHandler);
        document.removeEventListener('focusin', focusHandler);
      };
    }, [items, onChange]);

    const selectToolbarMenuItem = React.useCallback(
      (text: string) => {
        if (typeof currentIndex === 'number' && currentIndex < items.length) {
          const updatedItems = [...items];

          if (typeof calloutTargetElement?.selectionStart === 'number') {
            const item = updatedItems[currentIndex];
            const start = calloutTargetElement.selectionStart;
            const end =
              typeof calloutTargetElement?.selectionEnd === 'number'
                ? calloutTargetElement.selectionEnd
                : calloutTargetElement.selectionStart;
            const context = getCursorContextWithinLine(item.substring(0, start));
            const insertText = context === 'expression' ? text : `\${${text}}`;
            updatedItems[currentIndex] = [item.slice(0, start), insertText, item.slice(end)].join('');
            onChange(updatedItems);

            setTimeout(() => {
              calloutTargetElement.setSelectionRange(
                updatedItems[currentIndex].length,
                updatedItems[currentIndex].length
              );
            }, 0);
          }

          calloutTargetElement?.focus();
        }
      },
      [calloutTargetElement, currentIndex, items, onChange]
    );

    const insertSSMLTag = React.useCallback(
      (ssmlTagType: SSMLTagType) => {
        if (typeof currentIndex === 'number' && currentIndex < items.length) {
          const updatedItems = [...items];

          if (
            typeof calloutTargetElement?.selectionStart === 'number' &&
            typeof calloutTargetElement?.selectionEnd === 'number'
          ) {
            const item = updatedItems[currentIndex];
            const start = calloutTargetElement.selectionStart;
            const end = calloutTargetElement.selectionEnd;

            if (ssmlTagType === 'break' || ssmlTagType === 'audio') {
              const item = updatedItems[currentIndex];
              const start = calloutTargetElement.selectionStart;
              const end =
                typeof calloutTargetElement?.selectionEnd === 'number'
                  ? calloutTargetElement.selectionEnd
                  : calloutTargetElement.selectionStart;
              updatedItems[currentIndex] = [
                item.slice(0, start),
                `<${ssmlTagType} ${getSSMLProps(ssmlTagType)}/>`,
                item.slice(end),
              ].join('');
            } else {
              updatedItems[currentIndex] = [
                item.slice(0, start),
                `<${ssmlTagType} ${getSSMLProps(ssmlTagType)}>`,
                item.slice(start, end),
                `</${ssmlTagType}>`,
                item.slice(end),
              ].join('');
            }
            onChange(updatedItems);

            setTimeout(() => {
              calloutTargetElement.setSelectionRange(
                updatedItems[currentIndex].length,
                updatedItems[currentIndex].length
              );
            }, 0);
          }

          calloutTargetElement?.focus();
        }
      },
      [calloutTargetElement, currentIndex, items, onChange]
    );

    const toolbar = React.useMemo(
      () =>
        selectedKey === 'speak' ? (
          <LgSpeakModalityToolbar
            key="lg-speak-toolbar"
            lgTemplates={lgTemplates}
            properties={memoryVariables}
            onInsertSSMLTag={insertSSMLTag}
            onSelectToolbarMenuItem={selectToolbarMenuItem}
          />
        ) : (
          <LgEditorToolbar
            key="lg-toolbar"
            lgTemplates={lgTemplates}
            properties={memoryVariables}
            onSelectToolbarMenuItem={selectToolbarMenuItem}
          />
        ),
      [selectedKey, lgTemplates, memoryVariables, insertSSMLTag, selectToolbarMenuItem]
    );

    return (
      <div ref={containerRef}>
        {items.map((value, key) => (
          <StringArrayItem
            key={key}
            mode={key === currentIndex ? 'edit' : 'view'}
            value={value}
            onChange={handleChange(key)}
            onFocus={handleFocus(key)}
            onRemove={handleRemove(key)}
            onShowCallout={handleShowCallout}
          />
        ))}
        {currentIndex === null && (
          <Link as="button" styles={styles.link} onClick={handleClickAddVariation}>
            {formatMessage('Add new variation')}
          </Link>
        )}
        {calloutTargetElement && (
          <Callout
            directionalHint={DirectionalHint.topLeftEdge}
            gapSpace={2}
            isBeakVisible={false}
            target={calloutTargetElement}
          >
            {toolbar}
          </Callout>
        )}
      </div>
    );
  }
);

export { StringArrayEditor };
