import { notFound } from 'next/navigation';
import { prisma } from '@/lib/prisma';
import { HOME_CSS } from '../../home-css';
import { PROJECT_DETAIL_CSS } from '../project-detail-css';
import { H2X_LOGO } from '../../h2x-logo';
import SiteHeader from '../../site-header';
import NewsletterForm from '../../preview/newsletter-form';
import TrackPageView from '../../preview/track-page-view';

export const dynamic = 'force-dynamic';

const DEFAULT_NAV = [
  { label: 'Projects', href: '/projects' },
  { label: 'Studio', href: '/studio' },
  { label: 'Perspectives', href: '/perspectives' },
  { label: 'Practice', href: '/practice' },
  { label: 'Specialist', href: '/specialist' },
  { label: 'Ecosystem', href: '/ecosystem' },
  { label: 'Contact', href: '/contact' },
];

export default async function ProjectDetailPage({ params }: { params: { slug: string } }) {
  const project = await prisma.project.findUnique({ where: { slug: params.slug } });
  if (!project || !project.published) notFound();

  const [navItemsRaw, related] = await Promise.all([
    prisma.navItem.findMany({ where: { isActive: true }, orderBy: { order: 'asc' } }),
    prisma.project.findMany({
      where: { category: project.category, published: true, NOT: { id: project.id } },
      orderBy: { sortOrder: 'asc' },
      take: 4,
    }),
  ]);

  const navLinks = navItemsRaw.length > 0
    ? navItemsRaw.map((n) => ({ label: n.labelEn, href: n.href }))
    : DEFAULT_NAV;

  const paragraphs = (project.description || '').split('\n').filter(Boolean);

  return (
    <div className="h2x-site">
      <style dangerouslySetInnerHTML={{ __html: HOME_CSS }} />
      <style dangerouslySetInnerHTML={{ __html: PROJECT_DETAIL_CSS }} />
      <TrackPageView />

      <SiteHeader navLinks={navLinks} />

      <header className="hero" id="heroSection">
        <div className="hero-video" style={{ background: `url('${project.imageUrl}') center/cover no-repeat` }}></div>
        <div className="hero-scrim"></div>
        <div className="hero-inner">
          <h1>{project.name}</h1>
          <div className="hero-tags">
            <span>{project.type}</span>
            <span>{project.location}</span>
            <span>Completed <strong>{project.year}</strong></span>
          </div>
        </div>
      </header>

      <section className="project-info">
        <div className="wrap">
          <div className="project-info-grid">
            <div>
              {paragraphs.length > 0
                ? paragraphs.map((p, i) => <p key={i}>{p}</p>)
                : <p style={{ color: 'var(--stone-dark)' }}>Chưa có mô tả chi tiết cho dự án này — thêm trong <code>/admin/projects</code>.</p>}
            </div>
            <div>
              <div className="info-table">
                {project.client && <div className="info-row"><div className="label">Client</div><div className="value">{project.client}</div></div>}
                <div className="info-row"><div className="label">Location</div><div className="value">{project.location}</div></div>
                <div className="info-row"><div className="label">Team</div><div className="value">{project.team || 'H2X Studio'}</div></div>
                {project.expertise && <div className="info-row"><div className="label">Expertise</div><div className="value">{project.expertise}</div></div>}
                <div className="info-row"><div className="label">Completion</div><div className="value">{project.year}</div></div>
                {project.sector && <div className="info-row"><div className="label">Sector</div><div className="value">{project.sector}</div></div>}
                {project.photographer && <div className="info-row"><div className="label">Photographer</div><div className="value">{project.photographer}</div></div>}
              </div>
            </div>
          </div>
        </div>
      </section>

      {related.length > 0 && (
        <section className="work" style={{ paddingTop: 0, paddingBottom: 120 }}>
          <div className="wrap">
            <div className="section-head">
              <div>
                <div className="eyebrow">Related Projects</div>
                <h2>More in {project.category}.</h2>
              </div>
            </div>
            <div className="related-scroll">
              {related.map((p) => (
                <a key={p.id} href={`/projects/${p.slug}`} className="project-tile" style={{ flex: '0 0 320px' }}>
                  <div className="img">
                    <div style={{ width: '100%', height: '100%', background: `url('${p.imageUrl}') center/cover no-repeat` }}></div>
                  </div>
                  <div className="project-tile-tags">
                    <span>{p.type}</span><span>{p.location}</span><span>{p.year}</span>
                  </div>
                  <h3>{p.name}</h3>
                </a>
              ))}
            </div>
          </div>
        </section>
      )}

      <footer className="contact" id="contact">
        <div className="wrap">
          <div className="contact-top">
            <div className="eyebrow" style={{ color: 'var(--brass-light)', marginBottom: 18 }}>Enquire</div>
            <h2>Let&apos;s design something quiet, and considered.</h2>
            <a href="mailto:bd@h2xstudio.com" className="pill">Start a Project</a>
          </div>
          <div className="footer-grid">
            <div>
              <div className="footer-logo"><img src={H2X_LOGO} alt="H2X Studio" /></div>
              <p>Luxury hospitality interior design and build, based in Hanoi — working across Vietnam and Southeast Asia.</p>
              <div className="newsletter-block">
                <h4 style={{ fontSize: 11, letterSpacing: '.12em', textTransform: 'uppercase', color: 'var(--brass-light)', marginBottom: 14, fontWeight: 500 }}>Nhận tin từ H2X Studio</h4>
                <NewsletterForm />
              </div>
            </div>
            <div className="footer-col">
              <h4>Studio</h4>
              <a href="/projects">Projects</a>
              <a href="/studio#expertise">Expertise</a>
              <a href="/perspectives">Perspectives</a>
              <a href="/practice">Practice</a>
              <a href="/specialist">Specialist</a>
              <a href="/ecosystem">Ecosystem</a>
              <a href="/studio">About</a>
            </div>
            <div className="footer-col">
              <h4>Contact</h4>
              <a href="mailto:bd@h2xstudio.com">bd@h2xstudio.com</a>
              <span>0966 526 662</span>
              <span>26 Trần Hưng Đạo, Cửa Nam, Hoàn Kiếm, Hà Nội</span>
            </div>
            <div className="footer-col">
              <h4>Follow</h4>
              <a href="#">Instagram</a>
              <a href="#">Pinterest</a>
              <a href="#">YouTube</a>
            </div>
          </div>
          <div className="footer-bottom">
            <span>&copy; 2026 H2X Studio. All Rights Reserved</span>
          </div>
        </div>
      </footer>

      <div className="chat-widget">
        <a href="https://zalo.me/YOUR_ZALO_ID" target="_blank" rel="noopener noreferrer" className="chat-btn zalo" title="Chat qua Zalo">Zalo</a>
        <a href="https://wa.me/84YOUR_PHONE_NUMBER" target="_blank" rel="noopener noreferrer" className="chat-btn whatsapp" title="Chat qua WhatsApp">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38a9.9 9.9 0 0 0 4.74 1.21h.01c5.46 0 9.91-4.45 9.91-9.91C21.96 6.46 17.51 2 12.04 2Zm5.8 14.03c-.24.68-1.4 1.3-1.93 1.38-.5.08-1.13.11-1.82-.12-.42-.14-.96-.32-1.65-.62-2.9-1.25-4.8-4.17-4.94-4.36-.14-.19-1.18-1.57-1.18-3 0-1.42.75-2.12 1.02-2.41.27-.29.58-.36.78-.36.2 0 .39 0 .56.01.18.01.42-.07.66.5.24.58.83 1.99.9 2.13.07.14.12.31.02.5-.1.19-.15.31-.29.48-.14.17-.3.37-.43.5-.14.14-.29.29-.12.57.17.29.75 1.24 1.62 2.01 1.11.99 2.05 1.3 2.34 1.44.29.14.46.12.63-.07.17-.19.71-.83.9-1.11.19-.29.38-.24.63-.14.26.1 1.65.78 1.93.92.29.14.48.21.55.33.07.12.07.68-.17 1.36Z" /></svg>
        </a>
      </div>
    </div>
  );
}
