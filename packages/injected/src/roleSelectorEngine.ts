/**
 * Copyright (c) Microsoft Corporation.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { parseAttributeSelector } from '@isomorphic/selectorParser';
import { normalizeWhiteSpace } from '@isomorphic/stringUtils';

import { documentHasShadowRoots } from './selectorEvaluator';
import { beginAriaCaches, endAriaCaches, getAriaChecked, getAriaDisabled, getAriaExpanded, getAriaLevel, getAriaPressed, getAriaRole, getAriaSelected, getElementAccessibleDescription, getElementAccessibleNameText, isElementHiddenForAria, kAriaCheckedRoles, kAriaExpandedRoles, kAriaLevelRoles, kAriaPressedRoles, kAriaSelectedRoles } from './roleUtils';
import { matchesAttributePart } from './selectorUtils';

import type { AttributeSelectorOperator, AttributeSelectorPart } from '@isomorphic/selectorParser';
import type { SelectorEngine, SelectorRoot } from './selectorEngine';

type RoleEngineOptions = {
  role: string;
  description?: string | RegExp;
  descriptionOp?: Exclude<AttributeSelectorOperator, '<truthy>'>;
  descriptionExact?: boolean;
  name?: string | RegExp;
  nameOp?: Exclude<AttributeSelectorOperator, '<truthy>'>;
  nameExact?: boolean;
  checked?: boolean | 'mixed';
  pressed?: boolean | 'mixed';
  selected?: boolean;
  expanded?: boolean;
  level?: number;
  disabled?: boolean;
  includeHidden?: boolean;
};

const kSupportedAttributes = ['selected', 'checked', 'pressed', 'expanded', 'level', 'disabled', 'name', 'description', 'include-hidden'];
kSupportedAttributes.sort();

function validateSupportedRole(attr: string, roles: string[], role: string) {
  if (!roles.includes(role))
    throw new Error(`"${attr}" attribute is only supported for roles: ${roles.slice().sort().map(role => `"${role}"`).join(', ')}`);
}

function validateSupportedValues(attr: AttributeSelectorPart, values: any[]) {
  if (attr.op !== '<truthy>' && !values.includes(attr.value))
    throw new Error(`"${attr.name}" must be one of ${values.map(v => JSON.stringify(v)).join(', ')}`);
}

function validateSupportedOp(attr: AttributeSelectorPart, ops: AttributeSelectorOperator[]) {
  if (!ops.includes(attr.op))
    throw new Error(`"${attr.name}" does not support "${attr.op}" matcher`);
}

function validateAttributes(attrs: AttributeSelectorPart[], role: string): RoleEngineOptions {
  const options: RoleEngineOptions = { role };
  for (const attr of attrs) {
    switch (attr.name) {
      case 'checked': {
        validateSupportedRole(attr.name, kAriaCheckedRoles, role);
        validateSupportedValues(attr, [true, false, 'mixed']);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.checked = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      case 'pressed': {
        validateSupportedRole(attr.name, kAriaPressedRoles, role);
        validateSupportedValues(attr, [true, false, 'mixed']);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.pressed = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      case 'selected': {
        validateSupportedRole(attr.name, kAriaSelectedRoles, role);
        validateSupportedValues(attr, [true, false]);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.selected = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      case 'expanded': {
        validateSupportedRole(attr.name, kAriaExpandedRoles, role);
        validateSupportedValues(attr, [true, false]);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.expanded = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      case 'level': {
        validateSupportedRole(attr.name, kAriaLevelRoles, role);
        // Level is a number, convert it from string.
        if (typeof attr.value === 'string')
          attr.value = +attr.value;
        if (attr.op !== '=' || typeof attr.value !== 'number' || Number.isNaN(attr.value))
          throw new Error(`"level" attribute must be compared to a number`);
        options.level = attr.value;
        break;
      }
      case 'disabled': {
        validateSupportedValues(attr, [true, false]);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.disabled = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      case 'name': {
        if (attr.op === '<truthy>')
          throw new Error(`"name" attribute must have a value`);
        if (typeof attr.value !== 'string' && !(attr.value instanceof RegExp))
          throw new Error(`"name" attribute must be a string or a regular expression`);
        options.name = attr.value;
        options.nameOp = attr.op;
        options.nameExact = attr.caseSensitive;
        break;
      }
      case 'description': {
        if (attr.op === '<truthy>')
          throw new Error(`"description" attribute must have a value`);
        if (typeof attr.value !== 'string' && !(attr.value instanceof RegExp))
          throw new Error(`"description" attribute must be a string or a regular expression`);
        options.description = attr.value;
        options.descriptionOp = attr.op;
        options.descriptionExact = attr.caseSensitive;
        break;
      }
      case 'include-hidden': {
        validateSupportedValues(attr, [true, false]);
        validateSupportedOp(attr, ['<truthy>', '=']);
        options.includeHidden = attr.op === '<truthy>' ? true : attr.value;
        break;
      }
      default: {
        throw new Error(`Unknown attribute "${attr.name}", must be one of ${kSupportedAttributes.map(a => `"${a}"`).join(', ')}.`);
      }
    }
  }
  return options;
}

function queryRole(scope: SelectorRoot, options: RoleEngineOptions, internal: boolean): Element[] {
  const result: Element[] = [];
  const normalizedName = typeof options.name === 'string' ? normalizeWhiteSpace(options.name) : options.name;
  const nameOp = (internal && !options.nameExact && options.nameOp === '=') ? '*=' : (options.nameOp || '=');
  const nameAttrPart = options.name !== undefined ? { name: '', jsonPath: [], op: nameOp, value: normalizedName, caseSensitive: !!options.nameExact } : undefined;
  const normalizedDescription = typeof options.description === 'string' ? normalizeWhiteSpace(options.description) : options.description;
  const descriptionOp = (internal && !options.descriptionExact && options.descriptionOp === '=') ? '*=' : (options.descriptionOp || '=');
  const descriptionAttrPart = options.description !== undefined ? { name: '', jsonPath: [], op: descriptionOp, value: normalizedDescription, caseSensitive: !!options.descriptionExact } : undefined;
  const implicitTagSet = kImplicitTagsForRole[options.role];

  const match = (element: Element) => {
    if (implicitTagSet && !implicitTagSet.has(element.tagName) && !element.hasAttribute('role'))
      return;
    if (getAriaRole(element) !== options.role)
      return;
    if (options.selected !== undefined && getAriaSelected(element) !== options.selected)
      return;
    if (options.checked !== undefined && getAriaChecked(element) !== options.checked)
      return;
    if (options.pressed !== undefined && getAriaPressed(element) !== options.pressed)
      return;
    if (options.expanded !== undefined && getAriaExpanded(element) !== options.expanded)
      return;
    if (options.level !== undefined && getAriaLevel(element) !== options.level)
      return;
    if (options.disabled !== undefined && getAriaDisabled(element) !== options.disabled)
      return;
    if (!options.includeHidden) {
      const isHidden = isElementHiddenForAria(element);
      if (isHidden)
        return;
    }
    if (nameAttrPart) {
      // Always normalize whitespace in the accessible name.
      const accessibleName = normalizeWhiteSpace(getElementAccessibleNameText(element, !!options.includeHidden));
      if (!matchesAttributePart(accessibleName, nameAttrPart))
        return;
    }
    if (descriptionAttrPart) {
      // Always normalize whitespace in the accessible description.
      const accessibleDescription = normalizeWhiteSpace(getElementAccessibleDescription(element, !!options.includeHidden));
      if (!matchesAttributePart(accessibleDescription, descriptionAttrPart))
        return;
    }
    result.push(element);
  };

  const doc = scope.nodeType === 9 ? (scope as Document) : scope.ownerDocument;
  const mayHaveShadow = !doc || documentHasShadowRoots(doc);
  const fastSelector = (!mayHaveShadow && implicitTagSet) ? [...implicitTagSet].join(',') + ',[role]' : undefined;

  const query = (root: Element | ShadowRoot | Document) => {
    if (fastSelector) {
      for (const element of root.querySelectorAll(fastSelector))
        match(element);
      return;
    }
    const shadows: ShadowRoot[] = [];
    if ((root as Element).shadowRoot)
      shadows.push((root as Element).shadowRoot!);
    for (const element of root.querySelectorAll('*')) {
      match(element);
      if (element.shadowRoot)
        shadows.push(element.shadowRoot);
    }
    shadows.forEach(query);
  };

  query(scope);
  return result;
}

const kImplicitTagsForRole: Record<string, Set<string>> = {
  article: new Set(['ARTICLE']),
  banner: new Set(['HEADER']),
  blockquote: new Set(['BLOCKQUOTE']),
  button: new Set(['BUTTON', 'INPUT', 'SUMMARY']),
  caption: new Set(['CAPTION']),
  cell: new Set(['TD']),
  checkbox: new Set(['INPUT']),
  code: new Set(['CODE']),
  columnheader: new Set(['TH']),
  combobox: new Set(['SELECT', 'INPUT']),
  complementary: new Set(['ASIDE']),
  contentinfo: new Set(['FOOTER']),
  definition: new Set(['DD']),
  deletion: new Set(['DEL']),
  dialog: new Set(['DIALOG']),
  document: new Set(['HTML']),
  emphasis: new Set(['EM']),
  figure: new Set(['FIGURE']),
  form: new Set(['FORM']),
  gridcell: new Set(['TD']),
  group: new Set(['DETAILS', 'FIELDSET', 'OPTGROUP']),
  heading: new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6']),
  img: new Set(['IMG', 'SVG', 'svg']),
  insertion: new Set(['INS']),
  link: new Set(['A', 'AREA']),
  list: new Set(['UL', 'OL', 'MENU']),
  listbox: new Set(['SELECT', 'DATALIST']),
  listitem: new Set(['LI']),
  main: new Set(['MAIN']),
  mark: new Set(['MARK']),
  math: new Set(['MATH', 'math']),
  meter: new Set(['METER']),
  navigation: new Set(['NAV']),
  option: new Set(['OPTION']),
  paragraph: new Set(['P']),
  progressbar: new Set(['PROGRESS']),
  radio: new Set(['INPUT']),
  region: new Set(['SECTION']),
  row: new Set(['TR']),
  rowgroup: new Set(['TBODY', 'THEAD', 'TFOOT']),
  rowheader: new Set(['TH']),
  search: new Set(['SEARCH']),
  searchbox: new Set(['INPUT']),
  separator: new Set(['HR']),
  slider: new Set(['INPUT']),
  spinbutton: new Set(['INPUT']),
  status: new Set(['OUTPUT']),
  strong: new Set(['STRONG']),
  subscript: new Set(['SUB']),
  superscript: new Set(['SUP']),
  table: new Set(['TABLE']),
  term: new Set(['DFN', 'DT']),
  textbox: new Set(['INPUT', 'TEXTAREA']),
  time: new Set(['TIME']),
};

export function createRoleEngine(internal: boolean): SelectorEngine {
  return {
    queryAll: (scope: SelectorRoot, selector: string): Element[] => {
      const parsed = parseAttributeSelector(selector, true);
      const role = parsed.name.toLowerCase();
      if (!role)
        throw new Error(`Role must not be empty`);
      const options = validateAttributes(parsed.attributes, role);
      beginAriaCaches();
      try {
        return queryRole(scope, options, internal);
      } finally {
        endAriaCaches();
      }
    }
  };
}
