#pragma once
#include "rlTopologicalEntity2d.h"

class rlEdge2d :
	public rlTopologicalEntity2d
{
	friend class rlTopologicalEntSet;
public:
	rlSubClassTypeDef(rlTopologicalEntity2d,emEdge2d)

	void set_v1(rlId v);
	rlId v1() const;

	void set_v2(rlId v);
	rlId v2() const;

	void set_he12(rlId value);
	rlId he12() const;

	void set_he21(rlId value);
	rlId he21() const;

protected:
	rlEdge2d();
	~rlEdge2d();

private:
	rlId _v1;
	rlId _v2;

	rlId _he12;
	rlId _he21;
};

