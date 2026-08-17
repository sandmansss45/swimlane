import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react';

export interface IRegionFilterProps {
  // Distinct values present in the employee list's "region" field - as of
  // 2026-08-17 this is actually sourced from Department (e.g. "3101 -
  // Security and Safety"), since the real "QLE Existing Organisation" list
  // has no dedicated region column. Read from the live employee data
  // rather than hardcoded, so this adapts to whatever values are really
  // there either way.
  regions: string[];
  selectedRegion: string | undefined; // undefined = "All"
  onChange: (region: string | undefined) => void;
}

const ALL_KEY = '__all__';

// Narrows the employee picker before showing options - confirmed design
// rule (see the regions prop comment above for what "region" actually
// means for the real data).
const RegionFilter: React.FC<IRegionFilterProps> = ({ regions, selectedRegion, onChange }) => {
  return (
    <Pivot
      selectedKey={selectedRegion || ALL_KEY}
      onLinkClick={(item?: PivotItem) => {
        if (!item) return;
        onChange(item.props.itemKey === ALL_KEY ? undefined : item.props.itemKey);
      }}
    >
      <PivotItem headerText="All" itemKey={ALL_KEY} />
      {regions.map(region => (
        <PivotItem headerText={region} itemKey={region} key={region} />
      ))}
    </Pivot>
  );
};

export default RegionFilter;
