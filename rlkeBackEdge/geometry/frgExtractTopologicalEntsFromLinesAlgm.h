#pragma once
#pragma warning(disable : 4819)
#include <boost\geometry\index\rtree.hpp>
#include "rlTopologicalEnt2dSet.h"

class rlEdge2d;
class rlVertex2d;
class frgExtractTopologicalEntsFromLinesAlgm
{
public:
	typedef boost::geometry::model::point<double, 2, boost::geometry::cs::cartesian> Point_2d;
	typedef std::pair<Point_2d, rlId> Point2d_Id;
	typedef boost::geometry::model::box<Point_2d> Box;

	struct Vertex2dsOnSegment2d
	{
		AcGeLineSeg2d seg;
		std::map<double, rlVertex2d *, rl_double_sort1> vertex2ds;
	};

	frgExtractTopologicalEntsFromLinesAlgm(const AcDbObjectIdArray &line_ids);
	~frgExtractTopologicalEntsFromLinesAlgm();

	/** 返回值在外部释放 */
	rlTopologicalEnt2dSet *start();

private:
	frgExtractTopologicalEntsFromLinesAlgm();
	void _extract_vertices_from_lines(std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs, const AcDbObjectIdArray &ids);
	void _extract_from_seg(Vertex2dsOnSegment2d &stru);
	void _search_related_segs(AcDbObjectIdArray &ids, const AcGeLineSeg2d &seg);
	void _extract_vertices(Vertex2dsOnSegment2d &stru, const AcGeLineSeg2d &_seg);
	void _addto_rtree(std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs);
	void _addto_rtree(Vertex2dsOnSegment2d &stru);
	bool _is_in_rtree(rlId &id, const Point_2d &pnt);
	void _add_connections(const std::vector<Vertex2dsOnSegment2d> &seg_pnts_pairs);
	void _clear_inner_vertices_on_lines();

	bool _is_colinear_vertex(const rlVertex2d *v)const;
	bool _is_parallel_between(const rlEdge2d *e1, const rlEdge2d *e2) const;
	bool _remove_from_rtee(const Point_2d &pnt);

private:
	rlTopologicalEnt2dSet *_topologies;
	AcDbObjectIdArray _lines;
	boost::geometry::index::rtree<Point2d_Id, boost::geometry::index::rstar<16>> _rtree;
	resbuf *_filter;
};

