(() => {
  const filters = [...document.querySelectorAll('.publication-filter')];
  const publications = [...document.querySelectorAll('.publication-item')];
  const lists = [...document.querySelectorAll('.publication-list')];
  const yearGroups = [...document.querySelectorAll('.publication-year')];
  const emptyState = document.querySelector('.publication-empty');

  if (!filters.length || !publications.length) return;

  const applyFilter = (filter) => {
    let visibleCount = 0;

    filters.forEach((button) => {
      const active = button.dataset.filter === filter;
      button.classList.toggle('is-active', active);
      button.setAttribute('aria-pressed', String(active));
    });

    publications.forEach((publication) => {
      const tags = publication.dataset.tags.split(/\s+/).filter(Boolean);
      const visible = filter === 'all' || tags.includes(filter);
      publication.hidden = !visible;
      if (visible) visibleCount += 1;
    });

    lists.forEach((list) => {
      list.hidden = !list.querySelector('.publication-item:not([hidden])');
    });
    yearGroups.forEach((group) => {
      group.hidden = !group.querySelector('.publication-item:not([hidden])');
    });
    if (emptyState) emptyState.hidden = visibleCount > 0;
  };

  filters.forEach((button) => {
    button.addEventListener('click', () => applyFilter(button.dataset.filter));
  });

  applyFilter('all');
})();
