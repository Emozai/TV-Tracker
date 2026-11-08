import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  FlatList, Platform, Pressable, RefreshControl, SafeAreaView,
  ScrollView, StatusBar, StyleSheet, Text, TextInput, View,
} from 'react-native';

// ===== SETTINGS =====
// Your Apps Script web app URL (ends in /exec)
const API_URL =
  'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
// Must match API_KEY in Api.gs
const API_KEY = 'change-me-to-something-long';
// Column used to color the status badge (first header containing this word)
const STATUS_WORD = 'status';
// ====================

async function api(params) {
  const qs = new URLSearchParams({ ...params, key: API_KEY }).toString();
  const res = await fetch(`${API_URL}?${qs}`);
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch {
    throw new Error('Script did not return JSON. Check the doGet change and that access is "Anyone".');
  }
  if (!data.ok) throw new Error(data.error || 'Request failed');
  return data;
}

function statusColor(v) {
  const s = String(v || '').toLowerCase();
  if (/(done|complete|installed|ok|working|yes)/.test(s)) return C.good;
  if (/(progress|pending|partial|wip)/.test(s)) return C.warn;
  if (/(fail|fault|broken|no signal|missing|issue|not)/.test(s)) return C.bad;
  return C.muted;
}

// ---------- Screens ----------

function SheetsScreen({ onOpen }) {
  const [sheets, setSheets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [updated, setUpdated] = useState('');

  const load = useCallback(async () => {
    setLoading(true); setError('');
    try {
      const d = await api({ api: 'sheets' });
      setSheets(d.sheets); setUpdated(d.updated);
    } catch (e) { setError(e.message); }
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  return (
    <FlatList
      data={sheets}
      keyExtractor={(s) => s.name}
      contentContainerStyle={styles.list}
      refreshControl={<RefreshControl refreshing={loading} onRefresh={load} />}
      ListHeaderComponent={
        <View>
          <Text style={styles.h1}>Cabin TV</Text>
          <Text style={styles.sub}>PAX + Crew TV Tracker</Text>
          {!!error && <Text style={styles.error}>{error}</Text>}
          {!!updated && <Text style={styles.meta}>Updated {new Date(updated).toLocaleString()}</Text>}
        </View>
      }
      renderItem={({ item }) => (
        <Pressable style={styles.card} onPress={() => onOpen(item.name)}>
          <Text style={styles.cardTitle}>{item.name}</Text>
          <Text style={styles.count}>{item.rows} rows ›</Text>
        </Pressable>
      )}
    />
  );
}

function RowsScreen({ sheet, onBack, onOpenRow }) {
  const [data, setData] = useState({ headers: [], rows: [] });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [q, setQ] = useState('');

  const load = useCallback(async () => {
    setLoading(true); setError('');
    try { setData(await api({ api: 'rows', sheet })); }
    catch (e) { setError(e.message); }
    setLoading(false);
  }, [sheet]);

  useEffect(() => { load(); }, [load]);

  const statusKey = useMemo(
    () => data.headers.find((h) => h.toLowerCase().includes(STATUS_WORD)),
    [data.headers]);
  const titleKey = data.headers[0];
  const subKeys = data.headers.slice(1, 3).filter((h) => h !== statusKey);

  const filtered = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (!t) return data.rows;
    return data.rows.filter((r) =>
      data.headers.some((h) => String(r[h]).toLowerCase().includes(t)));
  }, [q, data]);

  return (
    <View style={{ flex: 1 }}>
      <View style={styles.topbar}>
        <Pressable onPress={onBack} hitSlop={12}><Text style={styles.back}>‹ Back</Text></Pressable>
        <Text style={styles.topTitle} numberOfLines={1}>{sheet}</Text>
        <View style={{ width: 50 }} />
      </View>
      <TextInput
        style={styles.search}
        placeholder="Search cabin, area, status…"
        placeholderTextColor={C.muted}
        value={q}
        onChangeText={setQ}
        autoCorrect={false}
      />
      {!!error && <Text style={[styles.error, { marginHorizontal: 16 }]}>{error}</Text>}
      <FlatList
        data={filtered}
        keyExtractor={(r) => String(r._row)}
        contentContainerStyle={styles.list}
        refreshControl={<RefreshControl refreshing={loading} onRefresh={load} />}
        ListHeaderComponent={!loading && <Text style={styles.meta}>{filtered.length} of {data.rows.length}</Text>}
        renderItem={({ item }) => (
          <Pressable style={styles.card} onPress={() => onOpenRow(item, data.headers)}>
            <View style={{ flex: 1 }}>
              <Text style={styles.cardTitle}>{item[titleKey] || `Row ${item._row}`}</Text>
              {subKeys.map((k) => !!item[k] && (
                <Text key={k} style={styles.cardSub} numberOfLines={1}>{k}: {item[k]}</Text>
              ))}
            </View>
            {statusKey && !!item[statusKey] && (
              <View style={[styles.badge, { backgroundColor: statusColor(item[statusKey]) }]}>
                <Text style={styles.badgeText}>{item[statusKey]}</Text>
              </View>
            )}
          </Pressable>
        )}
      />
    </View>
  );
}

function DetailScreen({ row, headers, onBack }) {
  return (
    <View style={{ flex: 1 }}>
      <View style={styles.topbar}>
        <Pressable onPress={onBack} hitSlop={12}><Text style={styles.back}>‹ Back</Text></Pressable>
        <Text style={styles.topTitle} numberOfLines={1}>{row[headers[0]] || `Row ${row._row}`}</Text>
        <View style={{ width: 50 }} />
      </View>
      <ScrollView contentContainerStyle={styles.list}>
        {headers.map((h) => (
          <View key={h} style={styles.field}>
            <Text style={styles.fieldLabel}>{h}</Text>
            <Text style={styles.fieldValue}>{row[h] || '—'}</Text>
          </View>
        ))}
        <Text style={styles.meta}>Sheet row {row._row}</Text>
      </ScrollView>
    </View>
  );
}

// ---------- App ----------

export default function App() {
  const [sheet, setSheet] = useState(null);
  const [detail, setDetail] = useState(null);

  return (
    <SafeAreaView style={styles.root}>
      <StatusBar barStyle="light-content" />
      {detail ? (
        <DetailScreen row={detail.row} headers={detail.headers} onBack={() => setDetail(null)} />
      ) : sheet ? (
        <RowsScreen sheet={sheet} onBack={() => setSheet(null)}
          onOpenRow={(row, headers) => setDetail({ row, headers })} />
      ) : (
        <SheetsScreen onOpen={setSheet} />
      )}
    </SafeAreaView>
  );
}

const C = {
  bg: '#0B1220', card: '#151E30', text: '#EAF0FA', muted: '#8A97AD',
  accent: '#3BA3FF', good: '#1F9D63', warn: '#C98A10', bad: '#C63D3D',
};

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: C.bg, paddingTop: Platform.OS === 'android' ? StatusBar.currentHeight : 0 },
  list: { padding: 16, paddingBottom: 40 },
  h1: { color: C.text, fontSize: 28, fontWeight: '700', marginTop: 8 },
  sub: { color: C.accent, fontSize: 15, marginBottom: 12 },
  meta: { color: C.muted, fontSize: 12, marginBottom: 10 },
  error: { color: '#FF8A8A', backgroundColor: '#3A1717', padding: 10, borderRadius: 8, marginBottom: 10 },
  card: {
    backgroundColor: C.card, borderRadius: 12, padding: 14, marginBottom: 10,
    flexDirection: 'row', alignItems: 'center',
  },
  cardTitle: { color: C.text, fontSize: 16, fontWeight: '600', flexShrink: 1 },
  cardSub: { color: C.muted, fontSize: 13, marginTop: 3 },
  count: { color: C.muted, marginLeft: 'auto' },
  badge: { borderRadius: 999, paddingHorizontal: 10, paddingVertical: 4, marginLeft: 10, maxWidth: 130 },
  badgeText: { color: '#fff', fontSize: 12, fontWeight: '600' },
  topbar: {
    flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between',
    paddingHorizontal: 16, paddingVertical: 12,
  },
  back: { color: C.accent, fontSize: 16, width: 50 },
  topTitle: { color: C.text, fontSize: 17, fontWeight: '600', flex: 1, textAlign: 'center' },
  search: {
    backgroundColor: C.card, color: C.text, borderRadius: 10, marginHorizontal: 16,
    paddingHorizontal: 14, paddingVertical: 10, fontSize: 15,
  },
  field: { backgroundColor: C.card, borderRadius: 10, padding: 12, marginBottom: 8 },
  fieldLabel: { color: C.muted, fontSize: 12, marginBottom: 4 },
  fieldValue: { color: C.text, fontSize: 15 },
});
