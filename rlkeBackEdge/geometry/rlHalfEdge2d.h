#pragma once
#include "rlEdge2d.h"

class rlHalfEdge2d :
	public rlTopologicalEntity2d
{
	friend class rlTopologicalEntSet;
public:
	rlSubClassTypeDef(rlTopologicalEntity2d, emHalfEdge2d)


	void set_owner(rlId edge);
	rlId owner() const;

	void set_is_positive(bool value);
	bool is_positive() const;

private:
	rlHalfEdge2d();
	~rlHalfEdge2d();

private:
	rlId _owner;
	/** true:1->2;false:2->1 */
	bool _is_positive;
};

