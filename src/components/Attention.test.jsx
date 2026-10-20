import { describe, expect, it, vi, afterEach } from 'vitest';
import { attentionGroups, attentionRows } from './Attention';

const task = (id, values = {}) => ({ id, title: id, status: 'In Progress', organization: 'NPSL (Calibration Lab)', ...values });
const projects = [
  { id: 'a', projectKey: 'a', title: 'Alpha', measurementArea: 'Pressure', tasks: [
    task('Zulu work', { dueDate: '2026-01-01', deferredDate: '2027-03-01', estimatedHours: 8, ownerName: 'Zoe' }),
    task('Alpha work', { dueDate: '2027-04-01', estimatedHours: 8, ownerName: 'Zoe' }),
    task('Completed work', { status: 'Complete', dueDate: '2020-01-01', estimatedHours: 999 }),
    task('Other organization', { organization: 'PMA 260B', estimatedHours: 999 }),
  ] },
  { id: 'b', projectKey: 'b', title: 'Beta', measurementArea: 'AC Voltage', tasks: [
    task('Beta work', { dueDate: '2026-10-01', estimatedHours: 12, ownerName: 'Alex', status: 'Blocked' }),
  ] },
  { id: 'c', projectKey: 'c', title: 'No deadline', tasks: [task('Undated work')] },
];
const rows = (sort = 'due', search = '') => attentionRows(projects, 'NPSL (Calibration Lab)', search, sort);

afterEach(() => vi.useRealTimers());

describe('attention views', () => {
  it.each([
    ['due', ['b', 'a', 'c']],
    ['hours', ['a', 'b', 'c']],
    ['title', ['a', 'b', 'c']],
    ['owner', ['b', 'a', 'c']],
    ['measurement-area', ['b', 'a', 'c']],
  ])('orders projects by %s using only matching open tasks', (sort, expected) => {
    expect(attentionGroups(rows(sort), 'project', sort).map(group => group.key)).toEqual(expected);
  });

  it('summarizes effective deadlines, overdue counts and total hours within the current filter', () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-10-16T12:00:00'));
    const groups = attentionGroups(rows(), 'project', 'due');
    expect(groups[0]).toMatchObject({ key: 'b', nextDue: '2026-10-01', overdueCount: 1, hours: 12 });
    expect(groups[1]).toMatchObject({ key: 'a', nextDue: '2027-03-01', overdueCount: 0, hours: 16 });
    expect(groups[2]).toMatchObject({ key: 'c', nextDue: '', overdueCount: 0, hours: 0 });
    const filtered = attentionGroups(rows('hours', 'Zulu'), 'project', 'hours');
    expect(filtered).toHaveLength(1);
    expect(filtered[0]).toMatchObject({ key: 'a', hours: 8, nextDue: '2027-03-01' });
    expect(filtered[0].items).toHaveLength(1);
    expect(attentionGroups(rows('due', 'no match'), 'project', 'due')).toEqual([]);
  });

  it('keeps task name ordering inside status and project groups without changing the available tasks', () => {
    const sorted = rows('title');
    const statuses = attentionGroups(sorted, 'status', 'title');
    expect(statuses.map(group => group.key)).toEqual(['Blocked', 'In Progress']);
    expect(statuses[1].items.map(({ task }) => task.title)).toEqual(['Alpha work', 'Undated work', 'Zulu work']);
    const grouped = attentionGroups(sorted, 'project', 'title');
    expect(grouped[0].items.map(({ task }) => task.title)).toEqual(['Alpha work', 'Zulu work']);
    expect(grouped.flatMap(group => group.items).map(row => row.task.id).sort()).toEqual(sorted.map(row => row.task.id).sort());
  });
});
