import { QueryTypes } from 'sequelize';
import { Sequelize } from 'sequelize-typescript';

/** A shipment row of the ERP query (only the columns these functions read or set). */
export interface ShipmentRow {
  Article_ID: string;
  Transport_Method: string;
  Factory_Name?: string;
  Customer_ID?: string;
  [column: string]: unknown;
}

/**
 * Normalizes item.Transport_Method (SEA / AIR / Land) of a shipment row of the ERP query and
 * returns its port of departure. Sets item.Transport_Method on the given item.
 * The port rules follow `factoryName` (default: the row's own Factory_Name).
 */
export const resolveShipmentTransportAndPort = (
  item: ShipmentRow,
  factoryName: string = item.Factory_Name!,
): string => {
  let portOfDeparture = '';
  if (
    item?.Article_ID.trim().toLowerCase() === 'SAMPLE SHOE'.trim().toLowerCase()
  ) {
    item.Transport_Method = 'AIR';
    portOfDeparture = 'SGN';
  } else {
    switch (item?.Transport_Method.trim().toLowerCase()) {
      case '11 SC'.trim().toLowerCase():
      case 'Ocean'.trim().toLowerCase():
      case 'Sea Air Collect'.trim().toLowerCase():
      case 'SC'.trim().toLowerCase():
      case 'Sea/Air'.trim().toLowerCase():
      case 'Partial'.trim().toLowerCase():
        item.Transport_Method = 'SEA';
        break;
      case '20 CC'.trim().toLowerCase():
      case 'AC'.trim().toLowerCase():
      case '10 AC'.trim().toLowerCase():
      case 'CC'.trim().toLowerCase():
      case 'Air PP'.trim().toLowerCase():
      case 'Air Express'.trim().toLowerCase():
      case 'AE'.trim().toLowerCase():
        item.Transport_Method = 'AIR';
        break;
      case 'RAIL'.trim().toLowerCase():
      case 'Rail Collect'.trim().toLowerCase():
      case 'Rail Colle'.trim().toLowerCase():
        item.Transport_Method = 'Land';
        break;
      default:
        item.Transport_Method = 'SEA';
        break;
    }

    switch (factoryName.trim().toLowerCase()) {
      case 'LYV'.trim().toLowerCase():
      case 'LVL'.trim().toLowerCase():
      case 'LHG'.trim().toLowerCase():
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'SEA'.trim().toLowerCase()
        ) {
          portOfDeparture = 'VNCLP';
        }
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'Land'.trim().toLowerCase()
        ) {
          portOfDeparture = 'SOTH';
        }
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'AIR'.trim().toLowerCase()
        ) {
          portOfDeparture = 'SGN';
        }
        break;
      case 'LYM'.trim().toLowerCase():
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'SEA'.trim().toLowerCase()
        ) {
          portOfDeparture = 'MMRGN';
        } else {
          portOfDeparture = 'RGN';
        }
        break;
      case 'LYF'.trim().toLowerCase():
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'SEA'.trim().toLowerCase()
        ) {
          portOfDeparture = 'IDSRG';
        } else {
          portOfDeparture = 'CGK';
        }
        break;
      default:
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'SEA'.trim().toLowerCase()
        ) {
          portOfDeparture = 'VNCLP';
        }
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'Land'.trim().toLowerCase()
        ) {
          portOfDeparture = 'SOTH';
        }
        if (
          item.Transport_Method.trim().toLowerCase() ===
          'AIR'.trim().toLowerCase()
        ) {
          portOfDeparture = 'SGN';
        }
        break;
    }
  }

  return portOfDeparture;
};

/**
 * Port of arrival of a shipment row: the PortCode saved in CMW_PortCode for its transport method and customer
 * ('N/A' when there is none). Call it after resolveShipmentTransportAndPort, which normalizes Transport_Method.
 */
export const findPortOfArrival = async (
  dbEIP: Sequelize | undefined,
  item: ShipmentRow,
): Promise<string> => {
  const queryPoA = `SELECT PortCode
                        FROM CMW_PortCode
                        WHERE TransportMethod LIKE ? AND CustomerNumber LIKE ?`;
  // `result` is undefined without dbEIP: reading result[0] then throws, as it always did
  const result = (await dbEIP?.query<{ PortCode: string }>(queryPoA, {
    type: QueryTypes.SELECT,
    replacements: [`%${item.Transport_Method}%`, `%${item.Customer_ID}%`],
  })) as { PortCode: string }[];
  return result[0]?.PortCode.trim().toUpperCase() || 'N/A';
};

/**
 * Adds Port_Of_Departure / Port_Of_Arrival to every shipment row. The port rules follow `factory`
 * (default: each row's own Factory_Name).
 * The port of arrival only depends on the transport method and the customer, so each distinct pair is
 * looked up once and shared by its rows (one query per row would queue thousands of queries on the pool).
 */
export const addPortsToRows = async (
  rows: ShipmentRow[],
  dbEIP: Sequelize | undefined,
  factory?: string,
) => {
  const portsOfDeparture = rows.map((item) =>
    resolveShipmentTransportAndPort(item, factory),
  );

  const arrivalLookups = new Map<string, Promise<string>>();
  const portOfArrivalOf = (item: ShipmentRow) => {
    // the two values findPortOfArrival puts in its query (String() keeps null and undefined apart)
    const key = JSON.stringify([
      String(item.Transport_Method),
      String(item.Customer_ID),
    ]);
    let lookup = arrivalLookups.get(key);
    if (!lookup) {
      lookup = findPortOfArrival(dbEIP, item);
      arrivalLookups.set(key, lookup);
    }
    return lookup;
  };

  return Promise.all(
    rows.map(async (item, index) => ({
      ...item,
      Transport_Method: item.Transport_Method,
      Port_Of_Departure: portsOfDeparture[index],
      Port_Of_Arrival: await portOfArrivalOf(item),
    })),
  );
};
