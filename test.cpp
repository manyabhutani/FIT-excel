#ifndef __PROGTEST__

#include <cstdlib>
#include <cstdio>
#include <cstring>
#include <cctype>
#include <climits>
#include <cfloat>
#include <cassert>
#include <cmath>
#include <iostream>
#include <sstream>
#include <fstream>
#include <iomanip>
#include <string>
#include <array>
#include <vector>
#include <list>
#include <set>
#include <map>
#include <stack>
#include <queue>
#include <unordered_set>
#include <unordered_map>
#include <memory>
#include <algorithm>
#include <functional>
#include <iterator>
#include <stdexcept>
#include <variant>
#include <optional>
#include <compare>
#include <charconv>
#include <span>
#include <utility>
#include "expression.h"

using namespace std::literals;
using CValue = std::variant<std::monostate, double, std::string>;

constexpr unsigned SPREADSHEET_CYCLIC_DEPS = 0x01;
constexpr unsigned SPREADSHEET_FUNCTIONS = 0x02;
constexpr unsigned SPREADSHEET_FILE_IO = 0x04;
constexpr unsigned SPREADSHEET_SPEED = 0x08;
constexpr unsigned SPREADSHEET_PARSER = 0x10;
#endif /* __PROGTEST__ */


//-------------------------------------------------------------START--------------------------------------------------------------------------------//
//Class to find Cell position
class CPos {
public:
    CPos(std::string_view str) {
        if (str.empty()) {
            throw std::invalid_argument("Empty Cell.");
        }
        size_t index = 0;
        while (index < str.size() && std::isalpha(str[index])) {
            column = column * 26 + (std::toupper(str[index]) - 'A' + 1);
            index++;
        }
        if (index == 0)
            throw std::invalid_argument("No column letters in cell identifier.");

        if (index == str.size())
            throw std::invalid_argument("No row digits in cell identifier.");




        row = 0;

        while (index < str.size() && std::isdigit(str[index])) {
            row = row * 10 + (str[index] - '0' );
            ++index;
        }

        if (index != str.size())
            throw std::invalid_argument("Invalid characters in cell identifier.");

        if (row < 0)
            throw std::invalid_argument("Row number must be a positive integer.");

    }

    int getRow() const {
        return row;
    }

    int getCol() const {
        return column;
    }



private:
    int row = 0;
    int column = 0;

};

class TreeNode;

//Class representing a cell in a spreadsheet
class Cell {
public:
    void setValue(const CValue &val);

    CValue getValue(){
        return value;
    }

    void setExpressionTree(std::shared_ptr<TreeNode> tree, const std::string& expr);

    CValue evaluate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>>& context);

    std::shared_ptr<Cell> clone() const;

    std::shared_ptr<TreeNode> getExpressionTree() const;

    std::string getExpressionString() const{
        return expressionString;
    }

private:
    CValue value;
    std::shared_ptr<TreeNode> expressionTree;
    std::string expressionString;

};


//Abstract Class representing a node of the Abstract Syntax Tree
class TreeNode {
public:
    virtual ~TreeNode() {}
    virtual CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &) const = 0;
    virtual std::shared_ptr<TreeNode> clone() const = 0;
    virtual std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const = 0;
    virtual std::string toString() const = 0;
    virtual std::set<std::pair<int, int>> getReferences() const = 0;
};

class AddNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    AddNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs) : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto lval = left->calculate(context);
        auto rval = right->calculate(context);

        if (std::holds_alternative<std::string>(lval) || std::holds_alternative<std::string>(rval)) {
            std::string leftStr = std::holds_alternative<std::string>(lval) ? std::get<std::string>(lval) : std::to_string(std::get<double>(lval));
            std::string rightStr = std::holds_alternative<std::string>(rval) ? std::get<std::string>(rval) : std::to_string(std::get<double>(rval));
            return leftStr + rightStr;
        }

        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) + std::get<double>(rval);
        }

        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<AddNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<AddNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }

    std::string toString() const override {
        return "(" + left->toString() + "+" + right->toString() + ")";
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto& rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }
};

class SubNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    SubNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto lval = left->calculate(context);
        auto rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) - std::get<double>(rval);
        }
        return std::monostate();
    }
    std::shared_ptr<TreeNode> clone() const override{
        return std::make_shared<SubNode>(left->clone(), right->clone());

    }
    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<SubNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return "(" + left->toString() + "-" + right->toString() + ")";
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto& rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class MulNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    MulNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs) : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto lval = left->calculate(context);
        auto rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) * std::get<double>(rval);
        }
        return std::monostate();
    }
    std::shared_ptr<TreeNode> clone() const override{
        return std::make_shared<MulNode>(left->clone(), right->clone());

    }
    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<MulNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  "(" + left->toString() + "*" + right->toString() + ")" ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto& rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }


};

class NegNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> operand;

public:
    NegNode(std::shared_ptr<TreeNode> op) : operand(op) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue operandValue = operand->calculate(context);
        if (std::holds_alternative<double>(operandValue)) {
            return -std::get<double>(operandValue);
        }
        return std::monostate();
    }
    std::shared_ptr<TreeNode> clone() const override{
        return std::make_shared<NegNode>(operand->clone());

    }
    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<NegNode>(operand->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  "-" + operand->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        return operand->getReferences();
    }

};

class PowerNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> base;
    std::shared_ptr<TreeNode> exponent;

public:
    PowerNode(std::shared_ptr<TreeNode> base, std::shared_ptr<TreeNode> exponent)
            : base(base), exponent(exponent) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto lval = base->calculate(context);
        auto rval = exponent->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            double baseValue = std::get<double>(lval);
            double exponentValue = std::get<double>(rval);
            return std::pow(baseValue, exponentValue);
        }
        return std::monostate();
    }
    std::shared_ptr<TreeNode> clone() const override{
        return std::make_shared<PowerNode>(base->clone(), exponent->clone());

    }
    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<PowerNode>(base->adjustReferences(rowOffset, colOffset), exponent->adjustReferences(rowOffset, colOffset));
    }

    std::string toString() const override {
        return  base->toString() + "^" + exponent->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = exponent->getReferences();
        const auto rightRefs = base->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }



};

class DivNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> numerator;
    std::shared_ptr<TreeNode> denominator;

public:
    DivNode(std::shared_ptr<TreeNode> numerator, std::shared_ptr<TreeNode> denominator): numerator(numerator), denominator(denominator) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto lval = numerator->calculate(context);
        auto rval = denominator->calculate(context);

        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            double denominatorValue = std::get<double>(rval);
            if (denominatorValue == 0) {
                return std::monostate();
            }return std::get<double>(lval) / denominatorValue;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<DivNode>(numerator->clone(), denominator->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<DivNode>(numerator->adjustReferences(rowOffset, colOffset), denominator->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return   "(" + numerator->toString() + "/" + denominator->toString() + ")";
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = numerator->getReferences();
        const auto& rightRefs = denominator->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class ValueNode : public TreeNode {
private:
    CValue value;

public:
    explicit ValueNode(CValue val) : value(val) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>>& context) const override {
        return value;
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<ValueNode>(value);
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return clone();
    }
    std::string toString() const override {
        if (std::holds_alternative<double>(value)) {
            return std::to_string(std::get<double>(value));
        } else if (std::holds_alternative<std::string>(value)) {
            return  std::get<std::string>(value);
        }
        return "";
    }
    std::set<std::pair<int, int>> getReferences() const override {
        return {};
    }
};

class ReferenceNode : public TreeNode {
private:
    std::pair<int, int> reference;
    bool isRowAbsolute;
    bool isColAbsolute;
    int originRow;
    int originCol;

public:
    ReferenceNode(int row, int col, bool rowAbs, bool colAbs, int origRow, int origCol)
            : reference(row, col), isRowAbsolute(rowAbs), isColAbsolute(colAbs),
              originRow(origRow), originCol(origCol) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        auto it = context.find({reference.first, reference.second});
        if (it != context.end() && it->second) {
            return it->second->evaluate(context);
        }
        return std::monostate();
    }
    std::shared_ptr<TreeNode> clone() const override{
        return std::make_shared<ReferenceNode>(reference.first,reference.second, isRowAbsolute, isColAbsolute , originRow,originCol);

    }
    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        int adjustedRow = isRowAbsolute ? reference.first : reference.first + rowOffset;
        int adjustedCol = isColAbsolute ? reference.second : reference.second + colOffset;
        return std::make_shared<ReferenceNode>(adjustedRow, adjustedCol, isRowAbsolute, isColAbsolute, originRow, originCol);
    }

    std::string toString() const override {
        return idToLabel(reference);
    }

    std::string idToLabel(std::pair<int, int> cellId) const {
        int col = cellId.second;
        int row = cellId.first;
        std::string label;
        while (col > 0) {
            int remainder = (col - 1) % 26;
            label = char('A' + remainder) + label;
            col = (col - 1) / 26;
        }
        label += std::to_string(row);
        if (isColAbsolute) label.insert(0, "$");
        if (isRowAbsolute) label.insert(label.begin() + (isColAbsolute ? 2 : 1), '$');

        return label;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        return {reference};
    }
};

class EqNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    EqNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) == std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) == std::get<std::string>(rval) ? 1.0 : 0.0;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<EqNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<EqNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + "==" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class LtNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    LtNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) < std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) < std::get<std::string>(rval) ? 1.0 : 0.0;
        }

        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<LtNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<LtNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + "<" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class LeNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    LeNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) <= std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) <= std::get<std::string>(rval) ? 1.0 : 0.0;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<LeNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<LeNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + "<=" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class GtNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    GtNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) > std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) > std::get<std::string>(rval) ? 1.0 : 0.0;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<GtNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<GtNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + ">" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class GeNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    GeNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(rhs) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) >= std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) >= std::get<std::string>(rval) ? 1.0 : 0.0;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<GeNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<GeNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + ">=" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }

};

class NeNode : public TreeNode {
private:
    std::shared_ptr<TreeNode> left;
    std::shared_ptr<TreeNode> right;

public:
    NeNode(std::shared_ptr<TreeNode> lhs, std::shared_ptr<TreeNode> rhs)
            : left(lhs), right(std::move(rhs)) {}

    CValue calculate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>> &context) const override {
        CValue lval = left->calculate(context);
        CValue rval = right->calculate(context);
        if (std::holds_alternative<double>(lval) && std::holds_alternative<double>(rval)) {
            return std::get<double>(lval) != std::get<double>(rval) ? 1.0 : 0.0;
        }
        else if (std::holds_alternative<std::string>(lval) && std::holds_alternative<std::string>(rval)) {
            return std::get<std::string>(lval) != std::get<std::string>(rval) ? 1.0 : 0.0;
        }
        return std::monostate();
    }

    std::shared_ptr<TreeNode> clone() const override {
        return std::make_shared<NeNode>(left->clone(), right->clone());
    }

    std::shared_ptr<TreeNode> adjustReferences(int rowOffset, int colOffset) const override {
        return std::make_shared<NeNode>(left->adjustReferences(rowOffset, colOffset), right->adjustReferences(rowOffset, colOffset));
    }
    std::string toString() const override {
        return  left->toString() + "!=" + right->toString() ;
    }
    std::set<std::pair<int, int>> getReferences() const override {
        std::set<std::pair<int, int>> refs = left->getReferences();
        const auto rightRefs = right->getReferences();
        refs.insert(rightRefs.begin(), rightRefs.end());
        return refs;
    }


};



//Class to build the expression
class TreeBuilder : public CExprBuilder {

private:
    std::stack<std::shared_ptr<TreeNode>> nodes;
public:

    void opAdd() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<AddNode>(left, right));

    }

    void valNumber(double val) override {
        nodes.push(std::make_shared<ValueNode>(val));

    }

    void valString(std::string val) override {
        nodes.push(std::make_shared<ValueNode>(val));
    }

    void opSub() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<SubNode>(left, right));
    }

    void opMul() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<MulNode>(left, right));


    }

    void opDiv() override {
        auto denominator = popNode();
        auto numerator = popNode();
        nodes.push(std::make_shared<DivNode>(numerator, denominator));
    }

    void opPow() override {
        auto exponent = popNode();
        auto base = popNode();
        nodes.push(std::make_shared<PowerNode>(base, exponent));
    }

    void opNeg() override {
        auto operand = popNode();
        nodes.push(std::make_shared<NegNode>(operand));
    }

    void opEq() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<EqNode>(left, right));
    }

    void opNe() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<NeNode>(left, right));
    }

    void opLt() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<LtNode>(left, right));
    }

    void opLe() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<LeNode>(left, right));
    }

    void opGt() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<GtNode>(left, right));
    }

    void opGe() override {
        auto right = popNode();
        auto left = popNode();
        nodes.push(std::make_shared<GeNode>(left, right));
    }

    void valReference(std::string val) override {
        bool isRowAbsolute = false;
        bool isColAbsolute = false;
        int row = 0, col = 0;
        long unsigned int idx = 0;

        if (val[idx] == '$') {
            isColAbsolute = true;
            idx++;
        }
        std::string colPart;
        while (std::isalpha(val[idx])) {
            colPart += std::toupper(val[idx++]);
        }
        for (char ch: colPart) {
            col = col * 26 + (ch - 'A' + 1);
        }
        if (idx < val.size() && val[idx] == '$') {
            isRowAbsolute = true;
            idx++;
        }

        std::string rowPart;
        while (idx < val.size() && std::isdigit(val[idx])) {
            rowPart += val[idx++];
        }
        row = static_cast<int>(std::stoul(rowPart));

        nodes.push(std::make_shared<ReferenceNode>(row, col, isRowAbsolute, isColAbsolute, originRow, originCol));
    }

    void valRange(std::string val) override {}

    void funcCall(std::string fnName, int paramCount) override {}

    std::shared_ptr<TreeNode> getRoot() const{

        return nodes.top();
    }

    void setOrigin(int row, int col) {
        originRow = row;
        originCol = col;
    }

private:
    std::shared_ptr<TreeNode> popNode() {
        auto node = nodes.top();
        nodes.pop();
        return node;
    }

    int originRow, originCol;
};

// Definition of various Cell Class methods
void Cell::setValue(const CValue &val) {
    value = val;
    expressionTree = nullptr;
}

void Cell::setExpressionTree(std::shared_ptr<TreeNode> tree, const std::string& expr) {
    expressionTree = std::move(tree);
    expressionString = expr;
    value = std::monostate();
}

CValue Cell::evaluate(const std::map<std::pair<int, int>, std::shared_ptr<Cell>>& context) {
    if (expressionTree) {
        return expressionTree->calculate(context);
    }
    return value;
}

std::shared_ptr<Cell> Cell::clone() const {
    auto newCell = std::make_shared<Cell>();
    newCell->value = value;
    if (expressionTree) {
        newCell->expressionTree = expressionTree->clone();
    }
    return newCell;
}

std::shared_ptr<TreeNode> Cell::getExpressionTree() const {
    return expressionTree;
}


// Class representing a Spreadsheet
class CSpreadsheet {
public:
    static unsigned capabilities() {
              return SPREADSHEET_CYCLIC_DEPS;
    }

    CSpreadsheet() = default;

    CSpreadsheet(const CSpreadsheet& other) {
        for (const auto& [key, cell] : other.cells) {
            cells[key] = cell->clone();
        }
    }

    CSpreadsheet(CSpreadsheet&& other) noexcept : cells(std::move(other.cells)) {}

    CSpreadsheet& operator=(const CSpreadsheet& other) {
        if (this == &other) return *this;

        cells.clear();
        for (const auto& [key, cell] : other.cells) {
            cells[key] = cell->clone();
        }
        return *this;
    }

    CSpreadsheet& operator=(CSpreadsheet&& other) noexcept {
        cells = std::move(other.cells);
        return *this;
    }

    bool setCell(const CPos &pos, const std::string &contents) {
        std::pair<int, int> key = {pos.getRow(), pos.getCol()};
        auto &cell = cells[key];
        if (!cell) {
            cell = std::make_shared<Cell>();
        }

        if (contents.empty()) {
            cell->setValue(std::monostate());
            return true;
        }

        if (contents[0] == '=') {
            TreeBuilder builder;
            builder.setOrigin(pos.getRow(), pos.getCol());
            try {
                parseExpression(contents, builder);
                cell->setExpressionTree(builder.getRoot(), contents);
            } catch (const std::exception &e) {
                cell->setExpressionTree(nullptr, "");
                return false;
            }
        } else {
            try {
                double num = std::stod(contents);
                cell->setValue(num);
            } catch (const std::invalid_argument &) {
                cell->setValue(contents);
            }
        }
        return true;
    }




    CValue getValue(CPos pos) {
        std::pair<int, int> key = {pos.getRow(), pos.getCol()};

        auto it = cells.find(key);
        if (it == cells.end()) {
            return std::monostate();
        }

        std::set<std::pair<int, int>> visited, recStack;
        if (detectCycle(key, visited, recStack)) {
            return std::monostate();
        }

        return it->second->evaluate(cells);
    }


    void copyRect(CPos dst, CPos src, int w = 1, int h = 1) {
        int srcRow = src.getRow();
        int srcCol = src.getCol();
        int dstRow = dst.getRow();
        int dstCol = dst.getCol();
        int rowOffset = dstRow - srcRow;
        int colOffset = dstCol - srcCol;

        std::map<std::pair<int, int>, std::shared_ptr<Cell>> tempStorage;

        for (int r = 0; r < h; ++r) {
            for (int c = 0; c < w; ++c) {
                std::pair<int, int> srcPos = {srcRow + r, srcCol + c};
                std::pair<int, int> dstPos = {dstRow + r, dstCol + c};

                auto srcIt = cells.find(srcPos);
                if (srcIt != cells.end()) {
                    tempStorage[dstPos] = srcIt->second->clone();
                    if (auto exprTree = tempStorage[dstPos]->getExpressionTree()) {
                        auto adjustedTree = exprTree->adjustReferences(rowOffset, colOffset);
                        std::string newExpr = adjustedTree->toString();

                        tempStorage[dstPos]->setExpressionTree(adjustedTree, "="+newExpr);
                    }
                } else {

                        tempStorage[dstPos] = std::make_shared<Cell>();
                        tempStorage[dstPos]->setValue(std::monostate());

                }
            }
        }

        for (const auto& [pos, cell] : tempStorage) {
            cells[pos] = cell;
        }
    }


    bool save(std::ostream &os) const {
        try {
            for (const auto &[key, cell]: cells) {
                int row = key.first;
                int col = key.second;

                if (cell->getExpressionTree()) {
                    std::string expr = cell->getExpressionString();
                    os << columnIndexToLabel(col) << "|" << std::to_string(row) << "|" << expr << std::endl;
                } else {
                    const CValue &value = cell->getValue();
                    os << columnIndexToLabel(col) << "|" << std::to_string(row) << "|";
                    if (std::holds_alternative<double>(value)) {
                        os << std::get<double>(value) << std::endl;
                    } else if (std::holds_alternative<std::string>(value)) {
                        os << std::get<std::string>(value) << std::endl;
                    } else if (std::holds_alternative<std::monostate>(value)) {
                        os << "" << std::endl;
                    }
                }
            }
            return true;
        }
        catch (...) {
            return false;
        }
    }


    bool load(std::istream &is) {
        try {
            this->cells.clear();
            std::string line;
            while (std::getline(is, line)) {
                if (line.empty()) {
                    continue;
                }

                std::istringstream iss(line);
                std::string columnId, rowId, value;
                if (!std::getline(iss, columnId, '|') ||
                    !std::getline(iss, rowId, '|') ||
                    !std::getline(iss, value)) {
                    return false;
                }

                try {
                    CPos pos(columnId + rowId);

                    setCell(pos, value);
                } catch (const std::exception &) {
                    return false;
                }
            }
            return true;
        } catch (...) {
            return false;
        }
    }


private:
    std::map<std::pair<int, int>, std::shared_ptr<Cell>> cells;

    std::string columnIndexToLabel(int col) const {
        std::string label;
        while (col > 0) {
            int remainder = (col - 1) % 26;
            label = char('A' + remainder) + label;
            col = (col - 1) / 26;
        }
        return label;
    }

    bool detectCycle(const std::pair<int, int>& cellId, std::set<std::pair<int, int>>& visited, std::set<std::pair<int, int>>& recStack) {
        if (recStack.find(cellId) != recStack.end())
            return true;
        if (visited.find(cellId) != visited.end())
            return false;
        visited.insert(cellId);
        recStack.insert(cellId);
        auto it = cells.find(cellId);
        if (it != cells.end() && it->second->getExpressionTree()) {
            auto refs = it->second->getExpressionTree()->getReferences();
            for (const auto& ref : refs) {
                if (detectCycle(ref, visited, recStack))
                    return true;
            }
        }
        recStack.erase(cellId);
        return false;
    }

};


#ifndef __PROGTEST__

bool valueMatch(const CValue &r,
                const CValue &s) {
    if (r.index() != s.index())
        return false;
    if (r.index() == 0)
        return true;
    if (r.index() == 2)
        return std::get<std::string>(r) == std::get<std::string>(s);
    if (std::isnan(std::get<double>(r)) && std::isnan(std::get<double>(s)))
        return true;
    if (std::isinf(std::get<double>(r)) && std::isinf(std::get<double>(s)))
        return (std::get<double>(r) < 0 && std::get<double>(s) < 0)
               || (std::get<double>(r) > 0 && std::get<double>(s) > 0);
    return fabs(std::get<double>(r) - std::get<double>(s)) <= 1e8 * DBL_EPSILON * fabs(std::get<double>(r));
}

int main() {

    CSpreadsheet x0, x1;
    std::ostringstream oss;
    std::istringstream iss;
    std::string data;

    assert (x0.setCell(CPos("A1"), "10"));
    assert (x0.setCell(CPos("A2"), "20.5"));
    assert (x0.setCell(CPos("A3"), "3e1"));
    assert (x0.setCell(CPos("A4"), "=40"));
    assert (x0.setCell(CPos("A5"), "=5e+1"));
    assert (x0.setCell(CPos("A6"), "raw text with any characters, including a quote \" or a newline\n"));
    assert (x0.setCell(CPos("A7"),"=\"quoted string, quotes must be doubled: \"\". Moreover, backslashes are needed for C++.\""));
    assert (valueMatch(x0.getValue(CPos("A1")), CValue(10.0)));
    assert (valueMatch(x0.getValue(CPos("A2")), CValue(20.5)));
    assert (valueMatch(x0.getValue(CPos("A3")), CValue(30.0)));
    assert (valueMatch(x0.getValue(CPos("A4")), CValue(40.0)));
    assert (valueMatch(x0.getValue(CPos("A5")), CValue(50.0)));
    assert (valueMatch(x0.getValue(CPos("A6")),
                       CValue("raw text with any characters, including a quote \" or a newline\n")));
    assert (valueMatch(x0.getValue(CPos("A7")),
                       CValue("quoted string, quotes must be doubled: \". Moreover, backslashes are needed for C++.")));
    assert (valueMatch(x0.getValue(CPos("A8")), CValue()));
    assert (valueMatch(x0.getValue(CPos("AAAA9999")), CValue()));


    assert (x0.setCell(CPos("B1"), "=A1+A2*A3"));
    assert (x0.setCell(CPos("B2"), "= -A1^2 -A2/2 "));
    assert (x0.setCell(CPos("B3"), "= 2 ^ $A$1"));
    assert (x0.setCell(CPos("B4"), "=($A1+A$2)^2"));
    assert (x0.setCell(CPos("B5"), "=B1+B2+B3+B4"));
    assert (x0.setCell(CPos("B6"), "=B1+B2+B3+B4+B5"));

    assert (valueMatch(x0.getValue(CPos("B1")), CValue(625.0)));

    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-110.25)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(1024.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(930.25)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(2469.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(4938.0)));


    assert (x0.setCell(CPos("A1"), "12"));
    assert (valueMatch(x0.getValue(CPos("B1")), CValue(627.0)));
    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-154.25)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(1056.25)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(5625.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(11250.0)));

    x1 = x0;
    assert (x0.setCell(CPos("A2"), "100"));
    assert (x1.setCell(CPos("A2"), "=A3+A5+A4"));
    assert (valueMatch(x0.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x0.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x0.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x0.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x0.getValue(CPos("B6")), CValue(38916.0)));

    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3612.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-204.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(17424.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(24928.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(49856.0)));

    oss.clear();
    oss.str("");
    assert (x0.save(oss));
    data = oss.str();
    iss.clear();
    iss.str(data);

    assert (x1.load(iss));
    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
    assert (x0.setCell(CPos("A3"), "4e1"));
    assert (valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
    assert (valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
    assert (valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
    assert (valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
    assert (valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
    assert (valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
    oss.clear();
    oss.str("");
    assert (x0.save(oss));
    data = oss.str();
    for (size_t i = 0; i < std::min<size_t>(data.length(), 10); i++)
        data[i] ^= 0x5a;
    iss.clear();
    iss.str(data);
    assert (!x1.load(iss));

    assert (x0.setCell(CPos("D0"), "10"));
    assert (x0.setCell(CPos("D1"), "20"));
    assert (x0.setCell(CPos("D2"), "30"));
    assert (x0.setCell(CPos("D3"), "40"));
    assert (x0.setCell(CPos("D4"), "50"));
    assert (x0.setCell(CPos("E0"), "60"));
    assert (x0.setCell(CPos("E1"), "70"));
    assert (x0.setCell(CPos("E2"), "80"));
    assert (x0.setCell(CPos("E3"), "90"));
    assert (x0.setCell(CPos("E4"), "100"));
    assert (x0.setCell(CPos("F10"), "=D0+5"));
    assert (x0.setCell(CPos("F11"), "=$D0+5"));
    assert (x0.setCell(CPos("F12"), "=D$0+5"));
    assert (x0.setCell(CPos("F13"), "=$D$0+5"));
    x0.copyRect(CPos("G11"), CPos("F10"), 1, 4);
    assert (valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F14")), CValue()));

    assert (valueMatch(x0.getValue(CPos("G10")), CValue()));

    assert (valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
//    std::cout<<std::get<double>(x0.getValue((CPos("G13"))));

    assert (valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
    assert (valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));

    x0.copyRect(CPos("G11"), CPos("F10"), 2, 4);
    assert (valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("F14")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G10")), CValue()));
    assert (valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
    assert (valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
    assert (valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));
    assert (valueMatch(x0.getValue(CPos("H10")), CValue()));

    assert (valueMatch(x0.getValue(CPos("H11")), CValue()));
    assert (valueMatch(x0.getValue(CPos("H12")), CValue()));
    assert (valueMatch(x0.getValue(CPos("H13")), CValue(35.0)));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue()));
    assert (x0.setCell(CPos("F0"), "-27"));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));
    x0.copyRect(CPos("H12"), CPos("H13"), 1, 2);
    assert (valueMatch(x0.getValue(CPos("H12")), CValue(25.0)));
    assert (valueMatch(x0.getValue(CPos("H13")), CValue(-22.0)));
    assert (valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));




    return EXIT_SUCCESS;



}

#endif /* __PROGTEST__ */
