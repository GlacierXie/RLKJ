#include "stdafx.h"
#include "frgExtractTopologicalEntsFromLinesAlgm.h"
#include "rlVertex2d.h"
#include "rlEdge2d.h"

frgExtractTopologicalEntsFromLinesAlgm::frgExtractTopologicalEntsFromLinesAlgm()
	: _topologies(new rlTopologicalEnt2dSet())
	, _filter(acutBuildList(-4, _T("<AND"), 8, _T("fmk_lines"), RTDXF0, _T("LINE"), -4, _T("AND>"), RTNONE))
{
}

frgExtractTopologicalEntsFromLinesAlgm::frgExtractTopologicalEntsFromLinesAlgm(const AcDbObjectIdArray &line_ids)
	: _topologies(new rlTopologicalEnt2dSet())
	, _lines(line_ids)
	, _filter(acutBuildList(-4, _T("<AND"), 8, _T("fmk_lines"), RTDXF0, _T("LINE"), -4, _T("AND>"), RTNONE))
{
}

frgExtractTopologicalEntsFromLinesAlgm::~frgExtractTopologicalEntsFromLinesAlgm()
{
	acutRelRb(_filter);
}

rlTopologicalEnt2dSet *frgExtractTopologicalEntsFromLinesAlgm::start()
{
	// - ����ÿһ���߶Σ��ҵ�ÿһ���߶������й����Ľڵ�
	std::vector<Vertex2dsOnSegment2d> seg_pnts_pairs;
	_extract_vertices_from_lines(seg_pnts_pairs, _lines);

	// - ���ӵ�rtree����
	_addto_rtree(seg_pnts_pairs);

	// - ����ÿһ���߶���ڵ�Ĺ�ϵ
	_add_connections(seg_pnts_pairs);

	// - �������ڵ�
	_clear_inner_vertices_on_lines();

	return _topologies;
}

void frgExtractTopologicalEntsFromLinesAlgm::_extract_vertices_from_lines(std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs,
	const AcDbObjectIdArray &ids)
{
	AcDbEntity *entity = NULL;
	acedSetStatusBarProgressMeter(_T("������ȡÿ���߶��ϵĽڵ�..."), 0, ids.length());
	for (int i = 0; i < ids.length(); i++)
	{
		acdbOpenAcDbEntity(entity, ids[i], AcDb::kForRead);
		if (entity == NULL)
			continue;

		if (entity->isA() != AcDbLine::desc())
		{
			entity->close();
			continue;
		}
		AcDbLine *line = (AcDbLine *)entity;

		Vertex2dsOnSegment2d stru;
		stru.seg.set(AcGePoint2d(line->startPoint().x, line->startPoint().y),
			AcGePoint2d(line->endPoint().x, line->endPoint().y));
		entity->close();

		_extract_from_seg(stru);
		seg_pnts_pairs.push_back(stru);
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgExtractTopologicalEntsFromLinesAlgm::_extract_from_seg(Vertex2dsOnSegment2d &stru)
{
	// - ��ȡ��ص��߶Σ��ཻ���غϣ�
	AcDbObjectIdArray ids;
	_search_related_segs(ids, stru.seg);

	// - ��Ϊ������д����ཻ���ص�
	// -- ���������յ�
	rlVertex2d *v = NULL;
	stru.vertex2ds.insert(std::make_pair(0.0, v));
	stru.vertex2ds.insert(std::make_pair(1.0, v));

	// -- ����ÿһ���߶�
	AcDbEntity *entity = NULL;
	for (int i = 0; i < ids.length(); i++)
	{
		acdbOpenAcDbEntity(entity, ids[i], AcDb::kForRead);
		if (entity == NULL)
			continue;

		if (entity->isA() != AcDbLine::desc())
		{
			entity->close();
			continue;
		}
		AcDbLine *_line = (AcDbLine *)entity;

		AcGeLineSeg2d _seg(AcGePoint2d(_line->startPoint().x, _line->startPoint().y),
			AcGePoint2d(_line->endPoint().x, _line->endPoint().y));
		_line->close();

		_extract_vertices(stru, _seg);
	}
}

void frgExtractTopologicalEntsFromLinesAlgm::_search_related_segs(AcDbObjectIdArray &ids, const AcGeLineSeg2d &seg)
{
	double tolerance = rlTolerance::equal_point();
	if (seg.length() < 2 * tolerance)
		return;

	// - ���췴ѡ����
	// -- �����ĸ���
	AcGePoint2d p1, p2, p3, p4;

	AcGeVector2d dir = seg.direction();
	AcGeVector2d offset = dir;
	offset.rotateBy(rlPi / 2.0);

	p1 = seg.startPoint(); p1.transformBy(-offset * tolerance);
	p2 = seg.endPoint(); p2.transformBy(-offset * tolerance);
	p3 = seg.endPoint(); p3.transformBy(offset * tolerance);
	p4 = seg.startPoint(); p4.transformBy(offset * tolerance);

	// -- �������
	ads_point _p1, _p2, _p3, _p4;
	_p1[0] = p1.x; _p1[1] = p1.y, _p1[2] = 0;
	_p2[0] = p2.x; _p2[1] = p2.y, _p2[2] = 0;
	_p3[0] = p3.x; _p3[1] = p3.y, _p3[2] = 0;
	_p4[0] = p4.x; _p4[1] = p4.y, _p4[2] = 0;
	resbuf *rect = acutBuildList(RTPOINT, _p1, RTPOINT, _p2, RTPOINT, _p3, RTPOINT, _p4, RTNONE);

	// - �����ѯ
	ads_name ss;
	int ret = acedSSGet(_T("CP"), rect, NULL, _filter, ss);
	acutRelRb(rect);
	if (ret != RTNORM)
		return;

	// - ��ȡ��ѯ���
	long len = 0;
	acedSSLength(ss, &len);
	for (int i = 0; i < len; i++)
	{
		ads_name name;
		acedSSName(ss, i, name);

		AcDbObjectId id;
		if (acdbGetObjectId(id, name) == Acad::eOk)
			ids.append(id);
	}
	acedSSFree(ss);
}

void frgExtractTopologicalEntsFromLinesAlgm::_extract_vertices(Vertex2dsOnSegment2d &stru, const AcGeLineSeg2d &_seg)
{
	// - �ж϶��ߵĹ�ϵ����_seg�ϵĵ㵽seg���ľ���Ϊkeyֵ��
	const AcGeLineSeg2d &seg = stru.seg;
	rlVertex2d *v = NULL;
	AcGePoint2d base = seg.startPoint();

	// -- �ж��Ƿ�ƽ��
	if (seg.isParallelTo(_seg, frgGlobals::Gtol))
	{
		if (seg.isOn(_seg.startPoint(), frgGlobals::Gtol))
			stru.vertex2ds.insert(std::make_pair(base.distanceTo(_seg.startPoint()), v));

		if (seg.isOn(_seg.endPoint(), frgGlobals::Gtol))
			stru.vertex2ds.insert(std::make_pair(base.distanceTo(_seg.endPoint()), v));

		return;
	}

	// -- �ж��ཻ
	AcGePoint2d pnt;
	if (seg.intersectWith(_seg, pnt, frgGlobals::Gtol))
		stru.vertex2ds.insert(std::make_pair(base.distanceTo(pnt), v));
}

void frgExtractTopologicalEntsFromLinesAlgm::_addto_rtree(std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs)
{
	acedSetStatusBarProgressMeter(_T("��������ÿ���ڵ㵽Rtree����..."), 0, (int)seg_pnts_pairs.size());
	for (int i = 0; i < seg_pnts_pairs.size(); i++)
	{
		Vertex2dsOnSegment2d &value = seg_pnts_pairs[i];
		_addto_rtree(value);
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgExtractTopologicalEntsFromLinesAlgm::_addto_rtree(Vertex2dsOnSegment2d &stru)
{
	const AcGeLineSeg2d &seg = stru.seg;
	std::map<double, rlVertex2d *, rl_double_sort1> &vertices = stru.vertex2ds;
	std::map<double, rlVertex2d *, rl_double_sort1>::iterator it = vertices.begin();

	AcGePoint2d base = seg.startPoint();
	AcGeVector2d dir = seg.direction();
	for (; it != vertices.end(); ++it)
	{
		AcGePoint2d pnt = base;
		pnt.transformBy(dir * it->first);

		Point_2d _pnt(pnt.x, pnt.y);
		rlId id = 0;
		if (_is_in_rtree(id, _pnt))
		{
			it->second = _topologies->get<rlVertex2d *>(id);
			assert(it->second);
		}
		else
		{
			rlVertex2d *_v = _topologies->_new<rlVertex2d>();
			_v->set_x(_pnt.get<0>());
			_v->set_y(_pnt.get<1>());

			Point2d_Id _pair(_pnt, _v->id());
			_rtree.insert(_pair);

			it->second = _v;
		}
	}
}

bool frgExtractTopologicalEntsFromLinesAlgm::_is_in_rtree(rlId &id, const Point_2d &pnt)
{
	double tol = rlTolerance::equal_point();
	Point_2d minPnt(pnt.get<0>() - tol, pnt.get<1>() - tol);
	Point_2d maxPnt(pnt.get<0>() + tol, pnt.get<1>() + tol);

	Box box(minPnt, maxPnt);
	std::vector<Point2d_Id> ret;
	_rtree.query(boost::geometry::index::intersects(box), std::back_inserter(ret));
	assert(ret.size() <= 1);

	if (ret.size() == 0)
		return false;

	id = ret[0].second;
	return true;
}

void frgExtractTopologicalEntsFromLinesAlgm::_add_connections(const std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs)
{
	acedSetStatusBarProgressMeter(_T("��������㡢�ߵ����˹�ϵ..."), 0, (int)seg_pnts_pairs.size());
	for (int i = 0; i < seg_pnts_pairs.size(); i++)
	{
		const Vertex2dsOnSegment2d &value = seg_pnts_pairs[i];
		if (value.vertex2ds.size() < 2)
			continue;
		std::map<double, rlVertex2d *, rl_double_sort1>::const_iterator it = value.vertex2ds.begin(), pos = it;
		pos++;
		for (; pos != value.vertex2ds.end(); ++pos, ++it)
		{
			rlVertex2d *v1 = it->second;
			rlVertex2d *v2 = pos->second;
			if (v1 == NULL || v2 == NULL)
			{
				assert(false);
				continue;
			}

			if (v1->has_coedge(v2))
				continue;

			rlEdge2d *_seg = _topologies->_new<rlEdge2d>();
			_seg->set_v1(v1->id());
			_seg->set_v2(v2->id());

			v1->insert(_seg->id());
			v2->insert(_seg->id());
		}
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgExtractTopologicalEntsFromLinesAlgm::_clear_inner_vertices_on_lines()
{
	// - ���ҷ��������ĵ�
	std::set<rlVertex2d *> filtered_vertices;

	const std::map<rlId, rlVertex2d*> &ptrs = _topologies->vertices();
	acedSetStatusBarProgressMeter(_T("���ڲ��Ҵ��ϲ��Ľڵ�..."), 0, (int)ptrs.size());
	std::map<rlId, rlVertex2d*>::const_iterator it = ptrs.begin();
	for (int i=0; it != ptrs.end(); ++it, i++)
	{
		rlVertex2d *v = it->second;
		if (_is_colinear_vertex(v) == false)
			continue;
		filtered_vertices.insert(v);
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();

	// - ����ÿһ���ڵ�
	acedSetStatusBarProgressMeter(_T("���ںϲ�ÿһ�����ҳ����Ľڵ�..."), 0, (int)filtered_vertices.size());
	std::set<rlVertex2d *>::const_iterator pos = filtered_vertices.begin();
	for (int i = 0; pos != filtered_vertices.end(); ++pos, i++)
	{
		rlVertex2d *v = *pos;

		// -- ��ȡ�ýڵ����ڵĽڵ�
		const std::set<rlId> &edges = v->edges();
		rlId eId1 = *edges.begin();
		rlId eId2 = *edges.rbegin();

		rlEdge2d *e1 = _topologies->get<rlEdge2d *>(eId1);
		rlEdge2d *e2 = _topologies->get<rlEdge2d *>(eId2);
		if (e1 == NULL || e2 == NULL)
		{
			assert(false);
			continue;
		}

		rlId vId1 = 0, vId2 = 0;
		if (v->has_another_end(vId1, e1) == false || v->has_another_end(vId2, e2) == false)
			continue;

		rlVertex2d *v1 = _topologies->get<rlVertex2d *>(vId1);
		rlVertex2d *v2 = _topologies->get<rlVertex2d *>(vId2);
		if (v1 == NULL || v2 == NULL)
		{
			assert(false);
			continue;
		}

		// -- ɾ�����еĹ�ϵ�������¹�ϵ
		// --- ɾ�����ڵ�������߶�id
		v1->erase(eId1);
		v2->erase(eId2);

		// --- ɾ���ڵ�
		_remove_from_rtee(Point_2d(v->x(), v->y()));
		_topologies->_delete(v);

		// --- ɾ�������߶�
		_topologies->_delete(e1);
		_topologies->_delete(e2);

		// --- �����߶�
		rlEdge2d *e = _topologies->_new<rlEdge2d>();
		e->set_v1(vId1);
		e->set_v2(vId2);

		v1->insert(e->id());
		v2->insert(e->id());
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

bool frgExtractTopologicalEntsFromLinesAlgm::_is_colinear_vertex(const rlVertex2d *v) const
{
	const std::set<rlId> &edges = v->edges();
	if (edges.size() != 2)
		return false;

	rlEdge2d *e1 = _topologies->get<rlEdge2d *>(*edges.begin());
	rlEdge2d *e2 = _topologies->get<rlEdge2d *>(*edges.rbegin());
	if (e1 == NULL || e2 == NULL)
	{
		assert(false);
		return false;
	}

	return _is_parallel_between(e1, e2);
}

bool frgExtractTopologicalEntsFromLinesAlgm::_is_parallel_between(const rlEdge2d *e1, const rlEdge2d *e2) const
{
	// - �����һ���߶�
	rlId vId1 = e1->v1();
	rlId vId2 = e1->v2();

	rlVertex2d *v1 = _topologies->get<rlVertex2d *>(vId1);
	rlVertex2d *v2 = _topologies->get<rlVertex2d *>(vId2);
	if (v1 == NULL || v2 == NULL)
	{
		assert(false);
		return false;
	}

	AcGeLineSeg2d seg1(AcGePoint2d(v1->x(), v1->y()), AcGePoint2d(v2->x(), v2->y()));

	// - ����ڶ����߶�
	vId1 = e2->v1();
	vId2 = e2->v2();

	v1 = _topologies->get<rlVertex2d *>(vId1);
	v2 = _topologies->get<rlVertex2d *>(vId2);
	if (v1 == NULL || v2 == NULL)
	{
		assert(false);
		return false;
	}

	AcGeLineSeg2d seg2(AcGePoint2d(v1->x(), v1->y()), AcGePoint2d(v2->x(), v2->y()));

	// - �ж϶����Ƿ�ƽ��
	if (seg1.isParallelTo(seg2, frgGlobals::Gtol) == Adesk::kTrue)
		return true;

	return false;
}

bool frgExtractTopologicalEntsFromLinesAlgm::_remove_from_rtee(const Point_2d &pnt)
{
	double tol = rlTolerance::equal_point();
	Point_2d minPnt(pnt.get<0>() - tol, pnt.get<1>() - tol);
	Point_2d maxPnt(pnt.get<0>() + tol, pnt.get<1>() + tol);

	Box box(minPnt, maxPnt);
	std::vector<Point2d_Id> ret;
	_rtree.query(boost::geometry::index::intersects(box), std::back_inserter(ret));
	if (ret.size() == 0)
		return false;

	for (int i = 0; i < ret.size(); i++)
		_rtree.remove(ret[i]);

	return true;
}