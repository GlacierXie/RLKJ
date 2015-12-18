#pragma once


class rlTopologicalEnt2dSet;
class rlVertex2d;
class rlEdge2d;
class rlHalfEdge2d;
class frgPolygon2dRecognitionAlgm
{
public:
	frgPolygon2dRecognitionAlgm(rlTopologicalEnt2dSet *value);
	~frgPolygon2dRecognitionAlgm();

	/** 集合在外部释放 */
	std::set<AcGePolyline2d *> start();

private:
	frgPolygon2dRecognitionAlgm();
	void _add_halfedges();
	void _search_vertices(std::set<AcGePolyline2d *> &out);
	void _remove_dead_vertices();
	void _edges_sort_on_vertex(std::map<double, rlEdge2d *> &sorted_edges, const rlVertex2d *v);
	void _get_sorted_edges(std::map<double, rlEdge2d *> &sorted_edges, rlVertex2d *v);
	bool _search_from(rlHalfEdge2d *edge, std::vector<rlHalfEdge2d *> &polygon);
	void _add_polygon(std::set<AcGePolyline2d *> &out, const AcGePoint2dArray &vertices);
	void _extract_pologon_vertices(AcGePoint2dArray &vertices, const std::vector<rlHalfEdge2d *> &polygon);

	rlHalfEdge2d *_get_forward_he(rlVertex2d *v, rlEdge2d *edge);
	rlHalfEdge2d *_get_next_he(rlHalfEdge2d *cHe);
	void _set_used_halfedges(const std::vector<rlHalfEdge2d *> &polygon);
	bool _polygon_is_counterclockwise(const AcGePoint2dArray &vertices);
private:
	rlTopologicalEnt2dSet *_topologies_2d;

	/** 每个节点上边按照与x轴夹角的排序 */
	std::map<rlVertex2d *, std::map<double, rlEdge2d *>> sorted_edges_on_v;
	/** 已经用过的半边 */
	std::set<rlId> _used_hes;
};

