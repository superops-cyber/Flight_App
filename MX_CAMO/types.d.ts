/**
 * Generated Schema (cleaned)
 *
 * This file contains a few domain-first interfaces used by the
 * STC/CAMO tooling: aircraft, STC, and document metadata.
 * Keep the old generated sheet interfaces separate if you still
 * need the raw sheet shapes; these are the typed models used
 * by the indexing & migration code.
 */

export interface IAircraft {
  id: string;                   // internal unique id (sheet or UUID)
  registration: string;         // e.g. "PT-ABC"
  model: string;                // aircraft model
  serialNumber?: string;        // airframe serial
  ownerId?: string;             // reference to owner record
  stcs?: string[];              // list of applied STC identifiers
  status?: 'active' | 'storage' | 'retired' | string;
  notes?: string;
}

export interface ISTC {
  id: string;                   // canonical STC id (e.g. "STC SA-1234" / "SA-1234")
  title?: string;               // human-friendly title
  holder?: string;              // company/person holding the STC
  modelsAffected?: string[];    // affected aircraft models
  issuedDate?: string;          // ISO date
  expirationDate?: string;      // ISO date if applicable
  files?: string[];             // array of document ids in the library
  notes?: string;
}

export type DocumentType =
  | 'Manual'
  | 'Certificate'
  | 'Drawing'
  | 'Supplement'
  | 'Form'
  | 'Approval'
  | 'Other';

export interface IDocument {
  id: string;                   // Drive file id
  name: string;                 // filename
  mimeType: string;
  md5Checksum?: string;         // for deduplication when available
  sizeBytes?: number;
  driveId?: string;             // shared drive id where found
  path?: string;                // best-effort full path in source drive
  stcIds?: string[];            // detected STC identifiers (may be empty)
  docType?: DocumentType;
  indexedAt?: string;           // ISO timestamp when indexed
  canonical?: boolean;          // true if this is the canonical copy in biblioteca
  notes?: string;
}

/* Keep a light mapping for sheet structure if you need it */
export interface ISheet_Structure {
  Aeronaves?: string;
  Description?: string;
}

/* Legacy generated sheet interfaces (kept minimal) */
export interface IAeronaves {
  AircraftID?: string | number;
  Registro?: string;
  Modelo?: string;
  OwnerID?: string | number;
  TT?: string | number;
  Status?: string;
}

export interface IProprietarios {
  OwnerID?: string | number;
  Nome?: string;
  Tipo?: string;
  Contato?: string;
}

export interface IComponentes {
  PN?: string;
  Numero_de_Serie?: string;
  Modelo?: string;
  Nome?: string;
  Categoria?: string;
  ProprietarioID?: string | number;
  Status?: string;
  Barcode?: string;
  Aviao?: string;
}

export interface IEventos_Manutencao {
  EventID?: string | number;
  ComponentID?: string | number;
  AircraftID?: string | number;
  TipoEvento?: string;
  Data?: string;
  TT_Aeronave?: string | number;
  Notas?: string;
  WO_ID?: string | number;
}

export interface IAD_SB {
  AD_Number?: string;
  Tipo?: string;
  ModeloAfetado?: string;
  IntervaloHoras?: string | number;
  ProximaData?: string;
  Status?: string;
}
  ClientID: any; // Edit 'any' to string, number, etc.
