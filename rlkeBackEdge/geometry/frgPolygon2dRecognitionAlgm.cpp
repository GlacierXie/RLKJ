#include "stdafx.h"
#include "frgPolygon2dRecognitionAlgm.h"
#include "rlTopologicalEnt2dSet.h"
#include "rlEdge2d.h"
#include "rlHalfEdge2d.h"
#include "rlVertex2d.h"
#include <boost/polygon/polygon.hpp>

typedef boost::polygon::polygon_data<double> B_Polygon;
typedef boost::polygon::polygon_traits<B_Polygon>::point_type B_Point;


frgPolygon2dRecognitionAlgm::frgPolygon2dRecognitionAlgm()
	: _topologies_2d(NULL)
{
}

frgPolygon2dRecognitionAlgm::frgPolygon2dRecognitionAlgm(rlTopologicalEnt2dSet *value)
	: _topologies_2d(value)
{

}


frgPolygon2dRecognitionAlgm::~frgPolygon2dRecognitionAlgm()
{
}

std::set<AcGePolyline2d *> frgPolygon2dRecognitionAlgm::start()
{
	std::set<AcGePolyline2d *> out;

	// - 前处理：清理端节点
	_remove_dead_vertices();

	// - 增加每个线段的半边
	_add_halfedges();

	// - 循环每一个节点进行查询
	_search_vertices(out);

	return out;
}

void frgPolygon2dRecognitionAlgm::_remove_dead_vertices()
{
	const std::map<rlId, rlVertex2d *> &vertices = _topologies_2d->vertices();

	// - 查找符合条件的点
	std::set<rlVertex2d *> will_removeds;
	std::map<rlId, rlVertex2d *>::const_iterator it = vertices.begin();
	acedSetStatusBarProgressMeter(_T("正在查找待清理的点..."), 0, (int)vertices.size());
	for (int i = 0; it != vertices.end(); ++it, i++)
	{
		rlVertex2d *v = it->second;
		if (v->edges().size() < 2)
			will_removeds.insert(v);
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();

	// - 循环每一个待删除的点进行处理
	std::set<rlVertex2d *>::const_iterator pos = will_removeds.begin();
	acedSetStatusBarProgressMeter(_T("正在清理堵塞的点和线..."), 0, (int)will_removeds.size());
	for (int i = 0; pos != will_removeds.end(); ++pos, i++)
	{
		rlVertex2d *v = *pos;

		// -- 判断是否为孤立点
		if (v->edges().size() == 0)
		{
			_topologies_2d->_delete(v);
			continue;
		}

		// -- 删除当前节点，并判断下一个节点是否也符合条件，并判断是否删除，直到不满足条件
		rlVertex2d *cur = v;
		while (true)
		{
			rlEdge2d *edge = _topologies_2d->get<rlEdge2d *>(*cur->edges().begin());

			// --- 获取相连的下一个节点
			rlId nId = 0;
			cur->has_another_end(nId, edge);
			rlVertex2d *next = _topologies_2d->get<rlVertex2d *>(nId);

			// --- 清理当前节点的关系
			next->erase(edge->id());
			_topologies_2d->_delete(edge);
			if (cur == v)
				_topologies_2d->_delete(cur);
			else if (will_removeds.find(cur) == will_removeds.end())
				_topologies_2d->_delete(cur);
			else
				break;

			// --- 判断是否循环下去
			if (next->edges().size() != 1)
				break;
			cur = next;
		}
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgPolygon2dRecognitionAlgm::_add_halfedges()
{
	// - 遍历每一个边
	const std::map<rlId, rlEdge2d *> &edges = _topologies_2d->edges();
	std::map<rlId, rlEdge2d *>::const_iterator it = edges.begin();
	acedSetStatusBarProgressMeter(_T("正在添加每个边的半边..."), 0, (int)edges.size());
	for (int i = 0; it != edges.end(); ++it, i++)
	{
		rlEdge2d *edge = it->second;

		rlHalfEdge2d *he12 = _topologies_2d->_new<rlHalfEdge2d>();
		edge->set_he12(he12->id());
		he12->set_owner(edge->id());

		rlHalfEdge2d *he21 = _topologies_2d->_new<rlHalfEdge2d>();
		he21->set_is_positive(false);
		edge->set_he21(he21->id());
		he21->set_owner(edge->id());

		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgPolygon2dRecognitionAlgm::_search_vertices(std::set<AcGePolyline2d *> &out)
{
	// - 遍历每一个节点
	const std::map<rlId, rlVertex2d *> &vertices = _topologies_2d->vertices();
	std::map<rlId, rlVertex2d *>::const_iterator it = vertices.begin();
	acedSetStatusBarProgressMeter(_T("正在循环每一个节点进行必和区域查找..."), 0, (int)vertices.size());
	for (int i = 0; it != vertices.end(); ++it, i++)
	{
		rlVertex2d *v = it->second;

		// -- 获取当前节点上排序的线集合，没有则新增
		std::map<double, rlEdge2d *> sorted_edges;
		_get_sorted_edges(sorted_edges, v);

		// -- 逆向循环每一个线段，以每个线段作为查找的起始
		std::map<double, rlEdge2d *>::const_reverse_iterator pos = sorted_edges.rbegin();
		for (; pos != sorted_edges.rend(); ++pos)
		{
			rlEdge2d *edge = pos->second;

			// --- 获取起始半边
			rlHalfEdge2d *he = NULL;
			he = _get_forward_he(v, edge);
			if (_used_hes.find(he->id()) != _used_hes.end())
				continue;

			// --- 以当前半边为起始，进行查询
			std::vector<rlHalfEdge2d *> polygon;
			bool ret = _search_from(he, polygon);
			if (ret == false)
				continue;

			// --- 提取多边形顶点
			AcGePoint2dArray _vertices;
			_extract_pologon_vertices(_vertices, polygon);

			// --- 判断多边形是否是逆时针
			bool is_counterclockwise = _polygon_is_counterclockwise(_vertices);

			// --- 增加多边形
			if (is_counterclockwise == false) //xzx false是最大外边框顺时针
				_add_polygon(out, _vertices);

			// --- 记录已走过的半边
			_set_used_halfedges(polygon);
		}
		acedSetStatusBarProgressMeterPos(i);
	}
	acedRestoreStatusBar();
}

void frgPolygon2dRecognitionAlgm::_get_sorted_edges(std::map<double, rlEdge2d *> &sorted_edges, rlVertex2d *v)
{
	std::map<rlVertex2d *, std::map<double, rlEdge2d *>>::const_iterator it = sorted_edges_on_v.find(v);
	if (it == sorted_edges_on_v.end())
		_edges_sort_on_vertex(sorted_edges, v);
	else
		sorted_edges = it->second;
}

void frgPolygon2dRecognitionAlgm::_edges_sort_on_vertex(std::map<double, rlEdge2d *> &sorted_edges, const rlVertex2d *v)
{
	const std::set<rlId> &edges = v->edges();
	std::set<rlId>::const_iterator it = edges.begin();

	for (; it != edges.end(); ++it)
	{
		rlEdge2d *edge = _topologies_2d->get<rlEdge2d *>(*it);

		rlVertex2d *v1 = _topologies_2d->get<rlVertex2d *>(edge->v1());
		rlVertex2d *v2 = _topologies_2d->get<rlVertex2d *>(edge->v2());

		AcGePoint2d sp, ep;
		if (v1 == v)
		{
			sp.set(v1->x(), v1->y());
			ep.set(v2->x(), v2->y());
		}
		else
		{
			ep.set(v1->x(), v1->y());
			sp.set(v2->x(), v2->y());
		}

		AcGeVector2d dir = ep - sp;
		dir.normalize();

		sorted_edges.insert(std::make_pair(dir.angle(), edge));
	}
}

bool frgPolygon2dRecognitionAlgm::_search_from(rlHalfEdge2d *sHe, std::vector<rlHalfEdge2d *> &polygon)
{
	bool finded = false;
	int m = 0;
	rlHalfEdge2d *cHe = sHe;
	while (true)
	{
		// - 加入半边到polygon
		polygon.push_back(cHe);

		// - 获取下一个半边
		rlHalfEdge2d *nHe = _get_next_he(cHe);
		if (nHe == NULL)
			break;

		// - 判断是否到达起始位置
		if (nHe == sHe)
		{
			finded = true;
			break;
		}

		// - 判断该半边是否走过
		if (_used_hes.find(nHe->id()) != _used_hes.end())
			break;

		cHe = nHe;
		if ((++m) > 100)
			break;
	}

	return finded;
}

void frgPolygon2dRecognitionAlgm::_add_polygon(std::set<AcGePolyline2d *> &out, const AcGePoint2dArray &vertices)
{
	AcGePolyline2d *_polygon = new AcGePolyline2d(vertices);
	out.insert(_polygon);
}

rlHalfEdge2d * frgPolygon2dRecognitionAlgm::_get_forward_he(rlVertex2d *v, rlEdge2d *edge)
{
	rlHalfEdge2d *he = NULL;

	if (edge->v1() == v->id())
		he = _topologies_2d->get<rlHalfEdge2d *>(edge->he12());
	else
		he = _topologies_2d->get<rlHalfEdge2d *>(edge->he21());

	return he;
}

rlHalfEdge2d *frgPolygon2dRecognitionAlgm::_get_next_he(rlHalfEdge2d *cHe)
{
	// - 获取所在边
	rlEdge2d *edge = _topologies_2d->get<rlEdge2d *>(cHe->owner());

	// - 获取下一个节点
	rlId nvId = cHe->is_positive() ? edge->v2() : edge->v1();
	rlVertex2d *nv = _topologies_2d->get<rlVertex2d *>(nvId);

	// - 获取当前节点上所有的排序的边
	std::map<double, rlEdge2d *> sorted_edges;
	_get_sorted_edges(sorted_edges, nv);

	// - 找到当前边
	std::map<double, rlEdge2d *>::const_reverse_iterator it = sorted_edges.rbegin();
	for (; it != sorted_edges.rend(); ++it)
	{
		if (it->second == edge)
			break;
	}
	if (it == sorted_edges.rend())
		return NULL;

	// - 反序找到下一个线段
	rlEdge2d *nedge = NULL;
	if ((++it) != sorted_edges.rend())
		nedge = it->second;
	else
		nedge = sorted_edges.rbegin()->second;

	return _get_forward_he(nv, nedge);
}

void frgPolygon2dRecognitionAlgm::_set_used_halfedges(const std::vector<rlHalfEdge2d *> &polygon)
{
	for (int i = 0; i < polygon.size(); i++)
	{
		_used_hes.insert(polygon[i]->id());
	}
}

void frgPolygon2dRecognitionAlgm::_extract_pologon_vertices(AcGePoint2dArray &vertices, const std::vector<rlHalfEdge2d *> &polygon)
{
	if (polygon.size() == 0)
		return;

	for (int i = 0; i < polygon.size(); i++)
	{
		rlHalfEdge2d *he = polygon[i];
		rlEdge2d *edge = _topologies_2d->get<rlEdge2d *>(he->owner());

		rlId svId = he->is_positive() ? edge->v1() : edge->v2();
		rlVertex2d *sv = _topologies_2d->get<rlVertex2d *>(svId);

		vertices.append(AcGePoint2d(sv->x(), sv->y()));
	}
	vertices.append(vertices[0]);
}

bool frgPolygon2dRecognitionAlgm::_polygon_is_counterclockwise(const AcGePoint2dArray &vertices)
{
	B_Point *pts = new B_Point[vertices.length()];
	for (int i = 0; i < vertices.length(); i++)
		pts[i] = boost::polygon::construct<B_Point>(vertices[i].x, vertices[i].y);

	B_Polygon polygon;
	boost::polygon::set_points(polygon, pts, pts + vertices.length());
	boost::polygon::direction_1d ret = boost::polygon::winding(polygon);
	SAFE_DELETE_ARR(pts);

	if (ret == boost::polygon::COUNTERCLOCKWISE)
		return true;
	return false;
}
