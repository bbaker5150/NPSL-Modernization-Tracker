export const workflowData = {
  phases: [
    {
      key: 'requirement',
      name: 'Requirement Phase',
      short: 'RP',
      acronym: 'RP',
      description: 'ICP and workload analysis.',
    },
    {
      key: 'development',
      name: 'Development Phase',
      short: 'DP',
      acronym: 'DP',
      description: 'Design, fabrication, integration, and technical development.',
    },
    {
      key: 'acquisition',
      name: 'Acquisition Phase',
      short: 'AP',
      acronym: 'AP',
      description: 'Purchasing activity after contract award.',
    },
    {
      key: 'operation-sustainment',
      name: 'Operation and Sustainment Phase',
      short: 'O&S',
      acronym: 'O&S',
      description: 'Transition following PMSD and METBUL publication.',
    },
    {
      key: 'production-procurement',
      name: 'Production/Procurement Phase',
      short: 'PP',
      acronym: 'PP',
      description: 'Bid testing or procurement of multiple parts.',
    },
  ],
  taskTemplates: [
    { title: 'Complete ICP and workload analysis', phaseKey: 'requirement', order: 1 },
    { title: 'Complete design, development, and technical documentation', phaseKey: 'development', order: 2 },
    { title: 'Complete acquisition and contract-award activities', phaseKey: 'acquisition', order: 3 },
    { title: 'Publish PMSD/METBUL and transition to sustainment', phaseKey: 'operation-sustainment', order: 4 },
    { title: 'Complete production/procurement and acceptance testing', phaseKey: 'production-procurement', order: 5 },
  ],
};
