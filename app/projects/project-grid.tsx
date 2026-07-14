'use client';

import { useState } from 'react';

type ProjectItem = {
  id: string;
  slug: string;
  name: string;
  category: string;
  type: string;
  location: string;
  year: string;
  imageUrl: string;
};

const FILTERS = [
  { key: 'all', label: 'All' },
  { key: 'hospitality', label: 'Hospitality' },
  { key: 'residential', label: 'Residential' },
  { key: 'fnb', label: 'F&B' },
  { key: 'special', label: 'Special Projects' },
];

export default function ProjectGrid({ projects }: { projects: ProjectItem[] }) {
  const [active, setActive] = useState('all');
  const filtered = active === 'all' ? projects : projects.filter((p) => p.category === active);

  return (
    <>
      <div className="filters" id="filters" style={{ marginBottom: 30 }}>
        {FILTERS.map((f) => (
          <button
            key={f.key}
            className={`filter-btn${active === f.key ? ' active' : ''}`}
            onClick={() => setActive(f.key)}
            type="button"
          >
            {f.label}
          </button>
        ))}
      </div>

      <div className="project-grid">
        {filtered.map((p) => (
          <a key={p.id} href={`/projects/${p.slug}`} className="project-tile" data-track-project={p.name}>
            <div className="img">
              <div style={{ width: '100%', height: '100%', background: `url('${p.imageUrl}') center/cover no-repeat` }}></div>
            </div>
            <div className="project-tile-tags">
              <span>{p.type}</span><span>{p.location}</span><span>{p.year}</span>
            </div>
            <h3>{p.name}</h3>
          </a>
        ))}
        {filtered.length === 0 && (
          <p style={{ color: 'var(--stone-dark)' }}>Không có dự án nào trong mục này.</p>
        )}
      </div>
    </>
  );
}
