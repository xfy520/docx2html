import { RelationshipTypes } from './types';

export const topLevelRels = [
  { type: RelationshipTypes.OfficeDocument, target: 'word/document.xml' },
  { type: RelationshipTypes.ExtendedProperties, target: 'docProps/app.xml' },
  { type: RelationshipTypes.CoreProperties, target: 'docProps/core.xml' },
  { type: RelationshipTypes.CustomProperties, target: 'docProps/custom.xml' },
];

export const A = 1;
